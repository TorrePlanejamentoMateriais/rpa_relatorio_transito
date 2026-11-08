# -*- coding: utf-8 -*-
"""Cruza FUP, Protocolos e Agendamentos. Requer: pip install openpyxl
Sem argumentos: abre seletores de arquivos. CLI: python gerar_fup.py --fup ...
--protocolos ... --agendamentos ... --saida ...
Mantem todas as linhas FUP e todos os protocolos distintos por pedido/item.
"""
import argparse
import re
import unicodedata
from collections import defaultdict, Counter
from copy import copy
from datetime import datetime, date
from decimal import Decimal, InvalidOperation
from pathlib import Path
import os
import tempfile
import traceback

from openpyxl import load_workbook, Workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter


def norm(value):
    text = unicodedata.normalize('NFKD', str(value or ''))
    return re.sub(r'[^a-z0-9]', '', ''.join(c for c in text if not unicodedata.combining(c)).lower())


def chave(value):
    if value is None:
        return ''
    text = str(value).strip()
    if not text:
        return ''
    if re.fullmatch(r'\d+(?:\.0+)?', text):
        return str(int(Decimal(text)))
    return text


def zero_dax(value):
    # DAX usa igualdade nao estrita: BLANK() = 0.
    # Texto invalido/erro NAO e convertido em zero.
    if value is None or (isinstance(value, str) and not value.strip()):
        return True
    if isinstance(value, bool):
        return False
    try:
        text = str(value).strip()
        if ',' in text:
            text = text.replace('.', '').replace(',', '.')
        return Decimal(text) == 0
    except (InvalidOperation, ValueError):
        return False


def status(ftf, valor, quantidade):
    return 'Pedido Entregue' if str(ftf or '').strip().upper() == 'X' or zero_dax(valor) or zero_dax(quantidade) else 'No Aguardo'


def localizar(wb, obrigatorias):
    for ws in wb:
        for cells in ws.iter_rows(min_row=1, max_row=20):
            values = [c.value for c in cells]
            mapa = {norm(v): i for i, v in enumerate(values) if v is not None}
            if all(norm(nome) in mapa for nome in obrigatorias):
                return ws, cells[0].row, mapa, values
    raise ValueError('Nao encontrei cabecalho com: ' + ', '.join(obrigatorias))


def ag_valor(value):
    if isinstance(value, (datetime, date)):
        return value
    if isinstance(value, (int, float)) and 1 <= value <= 100000:
        from openpyxl.utils.datetime import from_excel
        return from_excel(value)
    if value is None or not str(value).strip():
        return 'Sem data'
    text = str(value).strip()
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d/%m/%Y %H:%M:%S'):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            pass
    return text


def gerar(fup, protocolos, agendamentos, saida):
    fup, protocolos, agendamentos, saida = map(lambda p: Path(p).resolve(), (fup, protocolos, agendamentos, saida))
    if saida in (fup, protocolos, agendamentos):
        raise ValueError('A saida nao pode substituir nenhum arquivo de entrada.')
    if saida.suffix.lower() != '.xlsx':
        raise ValueError('A saida precisa terminar em .xlsx.')
    saida.parent.mkdir(parents=True, exist_ok=True)
    print('[1/4] Lendo agendamentos...', flush=True)
    agenda = defaultdict(list)
    wb = load_workbook(agendamentos, read_only=True, data_only=True)
    try:
        ws, linha, m, _ = localizar(wb, ['Nro. Protocolo', 'Ag Data'])
        for r in ws.iter_rows(min_row=linha + 1, values_only=True):
            p = chave(r[m[norm('Nro. Protocolo')]])
            if not p:
                continue
            v = ag_valor(r[m[norm('Ag Data')]])
            if v not in agenda[p]:
                agenda[p].append(v)
    finally:
        wb.close()
    print('[2/4] Indexando protocolos por Doc.compras + Item.1...', flush=True)
    indice = defaultdict(list)
    wb = load_workbook(protocolos, read_only=True, data_only=True)
    try:
        ws, linha, m, _ = localizar(wb, ['Nro. Protocolo', 'Doc.compras', 'Item.1'])
        for r in ws.iter_rows(min_row=linha + 1, values_only=True):
            doc = chave(r[m[norm('Doc.compras')]])
            item = chave(r[m[norm('Item.1')]])
            p = chave(r[m[norm('Nro. Protocolo')]])
            if doc and item and p and p not in indice[(doc, item)]:
                indice[(doc, item)].append(p)
    finally:
        wb.close()
    print('[3/4] Cruzando FUP e calculando status...', flush=True)
    origem = load_workbook(fup, read_only=True, data_only=True)
    destino = Workbook(write_only=True)
    cont = Counter()
    temp = None
    try:
        ws, linha, m, headers = localizar(origem, ['PO', 'Doc.compra', 'Itm', 'FtF', 'VAL PEND EF', 'QTD PEND na EM'])
        out = destino.create_sheet('FUP Atualizada')
        out.freeze_panes = 'A2'
        for letra, dim in getattr(ws, 'column_dimensions', {}).items():
            out.column_dimensions[letra].width = dim.width
        novos = ['Nro. Protocolo', 'Ag Data', 'Status']
        # Nomes novos nao substituem colunas preexistentes.
        usados = {norm(h) for h in headers}
        for i, nome in enumerate(novos):
            if norm(nome) in usados:
                novos[i] = nome + ' (Cruzamento)'
        cab = []
        for nome in headers + novos:
            c = WriteOnlyCell(out, value=nome)
            c.font = Font(bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='17365D')
            c.alignment = Alignment(wrap_text=True, vertical='center')
            cab.append(c)
        for i, largura in enumerate((22, 22, 23), len(headers) + 1):
            out.column_dimensions[get_column_letter(i)].width = largura
        out.append(cab)
        for cells in ws.iter_rows(min_row=linha + 1):
            vals = [c.value for c in cells]
            if all(v is None for v in vals):
                continue
            cont['Linhas FUP'] += 1
            doc = chave(vals[m[norm('Doc.compra')]])
            item = chave(vals[m[norm('Itm')]])
            # PO e a concatenacao Doc.compra + Itm. Tupla evita colisoes.
            ps = indice.get((doc, item), [])
            if not ps:
                cont['Linhas FUP sem protocolo'] += 1
            if len(ps) > 1:
                cont['Linhas FUP com varios protocolos'] += 1
            st = status(vals[m[norm('FtF')]], vals[m[norm('VAL PEND EF')]], vals[m[norm('QTD PEND na EM')]])
            cont['FUP: ' + st] += 1
            for p in ps or ['']:
                datas = agenda.get(p, []) if p else []
                if p and not datas:
                    cont['Protocolos vinculados sem agendamento'] += 1
                for data_ag in datas or ['Nao encontrado' if p else 'Sem protocolo']:
                    if cont['Linhas de saida'] >= 1048575:
                        raise ValueError('Resultado excede o limite de linhas do Excel.')
                    row = []
                    for original in cells:
                        c = WriteOnlyCell(out, value=original.value)
                        if getattr(original, 'has_style', False):
                            c.font = copy(original.font)
                            c.fill = copy(original.fill)
                            c.border = copy(original.border)
                            c.alignment = copy(original.alignment)
                            c.number_format = original.number_format
                        if original.data_type == 's':
                            c.data_type = 's'
                        row.append(c)
                    for v in (p, data_ag, st):
                        c = WriteOnlyCell(out, value=v)
                        if isinstance(v, (date, datetime)):
                            c.number_format = 'dd/mm/yyyy'
                        elif isinstance(v, str):
                            c.data_type = 's'
                        row.append(c)
                    row[-1].fill = PatternFill('solid', fgColor='E2F0D9' if st == 'Pedido Entregue' else 'FFF2CC')
                    out.append(row)
                    cont['Linhas de saida'] += 1
        out.auto_filter.ref = f'A1:{get_column_letter(len(headers)+3)}{cont["Linhas de saida"]+1}'
        resumo = destino.create_sheet('Resumo')
        resumo.append(['Indicador', 'Quantidade / Regra'])
        for k, v in cont.items():
            resumo.append([k, v])
        resumo.append(['Chave', 'FUP Doc.compra + Itm = Protocolos Doc.compras + Item.1'])
        resumo.append(['Multiplos resultados', 'Uma linha por protocolo e por Ag Data distintos; todas as linhas FUP mantidas.'])
        resumo.append(['Status', 'FtF = X OU VAL PEND EF = 0 OU QTD PEND na EM = 0. Vazios equivalem a zero, como no DAX.'])
        resumo.append(['Dados originais', 'Valores lidos do cache Excel. Salve e recalcule as entradas no Excel antes de executar.'])
        for nome, path in [('FUP', fup), ('Protocolos', protocolos), ('Agendamentos', agendamentos)]:
            resumo.append([nome, str(path)])
        resumo.column_dimensions['A'].width = 45
        resumo.column_dimensions['B'].width = 110
        print('[4/4] Salvando nova planilha...', flush=True)
        fd, temp = tempfile.mkstemp(suffix='.xlsx', dir=saida.parent)
        os.close(fd)
        destino.save(temp)
        os.replace(temp, saida)
        temp = None
    finally:
        origem.close()
        if temp and os.path.exists(temp):
            os.unlink(temp)
    print('Concluido:', saida, flush=True)
    for k, v in cont.items():
        print(f'  {k}: {v}')
    return dict(cont)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    for nome in ('fup', 'protocolos', 'agendamentos', 'saida'):
        parser.add_argument('--' + nome)
    args = parser.parse_args()
    root = None
    try:
        if not all((args.fup, args.protocolos, args.agendamentos, args.saida)):
            import tkinter as tk
            from tkinter import filedialog
            root = tk.Tk()
            root.withdraw()
            for nome in ('fup', 'protocolos', 'agendamentos'):
                if not getattr(args, nome):
                    path = filedialog.askopenfilename(title='Selecione: ' + nome.upper(), filetypes=[('Excel', '*.xlsx')])
                    if not path:
                        return
                    setattr(args, nome, path)
            if not args.saida:
                args.saida = filedialog.asksaveasfilename(title='Salvar nova FUP', initialfile='FUP_ATUALIZADA_' + datetime.now().strftime('%d.%m.%Y_%H%M%S') + '.xlsx', defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')])
                if not args.saida:
                    return
        gerar(args.fup, args.protocolos, args.agendamentos, args.saida)
        if root:
            from tkinter import messagebox
            messagebox.showinfo('Concluido', 'Nova FUP salva em:\n' + args.saida)
    except Exception as exc:
        traceback.print_exc()
        if root:
            from tkinter import messagebox
            messagebox.showerror('Erro', str(exc))
        raise
    finally:
        if root:
            root.destroy()


if __name__ == '__main__':
    main()
