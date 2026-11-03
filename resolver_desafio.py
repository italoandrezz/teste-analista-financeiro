"""Concilio lançamentos, apuro resultados e serializo fórmulas e caches OOXML."""
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
from datetime import datetime
from collections import Counter, defaultdict
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from copy import copy
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Side
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.properties import PageSetupProperties
from openpyxl.worksheet.views import Selection


def norm(value):
    return ' '.join(str(value or '').split()).casefold()


def cpf(value):
    return re.sub(r'\D', '', str(value or ''))


def money(value):
    text = str(value or 0).replace('R$', '').replace(' ', '').strip()
    if ',' in text:
        text = text.replace('.', '').replace(',', '.')
    return Decimal(text).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)


UNITS = ['Recife', 'Olinda', 'Caruaru', 'Afogados', 'Santa Cruz', 'Não classificado']
ALIASES = {norm(u): u for u in UNITS}
ALIASES.update({'sede/recife': 'Recife', 'afogados da ingazeira': 'Afogados',
                'santa cruz do capibaribe': 'Santa Cruz', 'sta cruz': 'Santa Cruz'})


def unit(value):
    if not norm(value):
        return 'Não classificado'
    if norm(value) not in ALIASES:
        raise ValueError(f'Unidade não mapeada: {value}')
    return ALIASES[norm(value)]


def cpf_formula(ref):
    expr = f'TRIM({ref}&"")'
    for char in ('.', '-', '/', ' '):
        expr = f'SUBSTITUTE({expr},"{char}","")'
    return '=' + expr


def unit_formula(ref):
    expr = '"REVISAR UNIDADE"'
    for alias, official in reversed(list(ALIASES.items())):
        expr = f'IF(LOWER(TRIM({ref}))="{alias}","{official}",{expr})'
    return f'=IF(TRIM({ref})="","Não classificado",{expr})'


def money_formula(ref):
    return f'=IF(ISNUMBER({ref}),ROUND({ref},2),ROUND(_xlfn.NUMBERVALUE(SUBSTITUTE(SUBSTITUTE({ref},"R$","")," ",""),",","."),2))'


def formatar_apresentacao(w):
    """Configuro estilos, dimensões e paginação das abas de apresentação."""
    answers, dre = w['RESPOSTAS'], w['RESULTADO_UNIDADE']
    for sheet, bounds, vertical in [(answers, (33, 8), 'top'), (dre, (10, 7), 'center')]:
        for row in sheet.iter_rows(max_row=bounds[0], max_col=bounds[1]):
            for cell in row:
                font = copy(cell.font)
                font.name, font.sz = 'Calibri', 11
                cell.font = font
                alignment = copy(cell.alignment)
                if alignment.horizontal == 'general':
                    alignment.horizontal = None
                alignment.vertical = vertical
                if sheet == answers or cell.row >= 3:
                    alignment.wrap_text = True
                cell.alignment = alignment
    for r in (1, 2, 5, 9, 12, 16, 18, 21, 23, 25, 32):
        # Copio o valor para a célula âncora antes de mesclar a linha.
        if answers.cell(r, 1).value is None:
            donor = next((answers.cell(r, c) for c in range(2, 9)
                          if answers.cell(r, c).value is not None), None)
            if donor is not None:
                answers.cell(r, 1).value = donor.value
                answers.cell(r, 1)._style = copy(donor._style)
        answers.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        if r == 5:
            border = copy(answers.cell(r, 1).border)
            border.top = Side(style='thin', color='FFBFBFBF')
            answers.cell(r, 1).border = border
        answers.cell(r, 1).alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    for col, width in {'A':6, 'B':66, 'C':19, 'D':29, 'E':19, 'F':25, 'G':55, 'H':7}.items():
        answers.column_dimensions[col].width = width + 0.7109375
    for r in range(6, 23):
        for col in ('D', 'F'):
            answers[f'{col}{r}'].alignment = Alignment(horizontal='right', vertical='top', wrap_text=True)
    for address in ('D15', 'F15', 'D17', 'D20', 'D10', 'D24', 'D26', 'D33'):
        answers[address].alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
    for r in (6, 7, 8, 11, 13, 14, 15, 17, 19, 20, 22):
        answers.row_dimensions[r].height = 78
    for r, height in {1:24, 2:24, 4:34, 10:84, 24:200, 26:66, 29:32.1, 30:46, 33:90}.items():
        answers.row_dimensions[r].height = height
    answers.page_margins.left = answers.page_margins.right = 18 / 72
    answers.page_margins.top = answers.page_margins.bottom = 24 / 72
    answers.print_title_rows = '1:4'
    dre.merge_cells('A1:G1')
    dre['A1'].alignment = Alignment(vertical='center', wrap_text=True)
    for address, label in {'B3':'Receita (R$)', 'C3':'(-) Custos (R$)',
                           'E3':'(-) Despesas (R$)', 'F3':'Resultado (R$)'}.items():
        dre[address] = label
    # Incluo o preenchimento de 0,7109375 na largura gravada no OOXML.
    dre.column_dimensions['A'].width = 24.7109375
    for col in ('B', 'C', 'E', 'F'):
        dre.column_dimensions[col].width = 21.7109375
        for r in range(4, 11):
            dre[f'{col}{r}'].number_format = '#,##0.00'
    for col in ('D', 'G'):
        dre.column_dimensions[col].width = 25.7109375
        for r in range(4, 11):
            dre[f'{col}{r}'].number_format = '0.0%'
    dre.row_dimensions[1].height = 42
    dre.row_dimensions[3].height = 48
    for r in range(4, 11):
        dre.row_dimensions[r].height = 30
        for col in range(1, 8):
            dre.cell(r, col).alignment = Alignment(horizontal='left' if col == 1 else 'right',
                                                  vertical='center', wrap_text=True)
    for sheet, last, pages, zoom in [(answers, 'H33', 2, 70), (dre, 'G10', 1, 85)]:
        # Separo as larguras agrupadas em B:C e E:G para evitar faixas sobrepostas.
        for c in range(1, 9 if sheet == answers else 8):
            dim = sheet.column_dimensions[openpyxl.utils.get_column_letter(c)]
            dim.min = dim.max = c
        sheet.page_setup.orientation = 'landscape'
        sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = pages
        sheet.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        sheet.print_area = f'A1:{last}'
        sheet.print_options.horizontalCentered = True
        sheet.sheet_view.zoomScale = sheet.sheet_view.zoomScaleNormal = zoom
        sheet.sheet_view.selection = [Selection(activeCell='A1', sqref='A1')]
    for sheet in w:
        sheet.sheet_view.tabSelected = sheet == answers
    w.active = w.sheetnames.index('RESPOSTAS')


def solve(source, destination):
    w = openpyxl.load_workbook(source)
    cache = {}
    def formula(sheet, cell, expression, value):
        sheet[cell] = expression
        cache[(sheet.title, cell)] = value
    classes = {
        'Honorários de Êxito': ('Receita', 'Remuneração pelo êxito dos serviços jurídicos.'),
        'Honorários Contratuais': ('Receita', 'Remuneração contratual dos serviços prestados.'),
        'Honorários de Sucumbência': ('Receita', 'Receita de honorários decorrente da sucumbência.'),
        'Consultas e Planejamento': ('Receita', 'Receita de consultas e planejamento previdenciário.'),
        'Comissões de Parceiros': ('Custo', 'Comissão diretamente vinculada à geração dos serviços/clientes.'),
        'Custas Processuais': ('Custo', 'Gasto diretamente identificável com a execução dos processos.'),
        'Laudos e Perícias': ('Custo', 'Serviço técnico diretamente aplicado aos processos dos clientes.'),
        'Pessoal Administrativo': ('Despesa', 'Estrutura administrativa de suporte ao escritório.'),
        'Marketing': ('Despesa', 'Divulgação e aquisição de clientes; função comercial.'),
        'Ocupação': ('Despesa', 'Manutenção da estrutura física das unidades.'),
        'Material e Serviços Gerais': ('Despesa', 'Suporte geral ao funcionamento do escritório.'),
        'Sistemas': ('Despesa', 'Infraestrutura tecnológica geral, sem vínculo individual com processo.'),
    }
    for r in range(2, 14):
        cat = w['CLASSIFICACAO'].cell(r, 1).value
        w['CLASSIFICACAO'].cell(r, 3, classes[cat][0])
        w['CLASSIFICACAO'].cell(r, 4, classes[cat][1])

    ps = w['BASE_PROCESSOS']
    processes = list(ps.iter_rows(min_row=2, max_col=9, values_only=True))
    counts = Counter(cpf(p[1]) for p in processes)
    assert max(counts.values()) == 1, 'CPF com vários processos: revisar regra de alocação.'
    products = {cpf(p[1]): p[4] for p in processes}
    ls = w['BASE_LANCAMENTOS']
    entries = list(ls.iter_rows(min_row=2, max_col=10, values_only=True))
    n, pn = len(entries) + 1, len(processes) + 1
    headers = ['CPF_CNPJ_Normalizado', 'Unidade_Padronizada', 'Valor_Numerico',
               'Valor_Original_Texto', 'Copia_Duplicada', 'Lancamento_Valido', 'Grupo',
               'Produto', 'Primeira_Ocorrencia_Linha']
    for col, header in enumerate(headers, 11):
        ls.cell(1, col, header)
    seen = {}
    totals = {u: defaultdict(Decimal) for u in UNITS}
    receipts = defaultdict(Decimal)
    receipt_count = Counter()
    product_totals = defaultdict(Decimal)
    duplicate_count = cancelled_count = text_count = valid_count = 0
    duplicate_amount = Decimal(0)
    for r, row in enumerate(entries, 2):
        ident, date, kind, category, branch, name, document, value, status, obs = row
        amount = money(value)
        key = (cpf(document), date, norm(category), amount)
        duplicate = key in seen
        first = seen.setdefault(key, r)
        valid = norm(status) == 'pago' and not duplicate
        group = classes[category][0]
        canonical = unit(branch)
        product = products.get(cpf(document), 'Sem processo correspondente') if category == 'Honorários de Êxito' else ''
        helper_values = [cpf(document), canonical, float(amount), int(isinstance(value, str)),
                         int(duplicate), int(valid), group, product, first]
        helper_formulas = [cpf_formula(f'G{r}'), unit_formula(f'E{r}'), money_formula(f'H{r}'),
            f'=--ISTEXT(H{r})',
            f'=--(COUNTIFS($K$2:K{r},K{r},$B$2:B{r},B{r},$D$2:D{r},D{r},$M$2:M{r},M{r})>1)',
            f'=--AND(TRIM(I{r})="Pago",O{r}=0)',
            f'=IFERROR(VLOOKUP(D{r},CLASSIFICACAO!$A$2:$C$13,3,FALSE),"REVISAR CATEGORIA")',
            f'=IF(D{r}="Honorários de Êxito",IFERROR(INDEX(BASE_PROCESSOS!$E$2:$E${pn},MATCH(K{r},BASE_PROCESSOS!$J$2:$J${pn},0)),"Sem processo correspondente"),"")',
            f'=MATCH(1,INDEX(($K$2:$K${n}=K{r})*($B$2:$B${n}=B{r})*($D$2:$D${n}=D{r})*($M$2:$M${n}=M{r}),0),0)+1']
        for c, (expr, cached_value) in enumerate(zip(helper_formulas, helper_values), 11):
            formula(ls, f'{openpyxl.utils.get_column_letter(c)}{r}', expr, cached_value)
        duplicate_count += duplicate
        cancelled_count += norm(status) == 'cancelado'
        text_count += isinstance(value, str)
        duplicate_amount += amount if duplicate else Decimal(0)
        if valid:
            assert datetime(2026, 7, 1) <= date <= datetime(2026, 9, 30)
            valid_count += 1
            totals[canonical][group] += amount
            if category == 'Honorários de Êxito':
                receipts[cpf(document)] += amount
                receipt_count[cpf(document)] += 1
                product_totals[product] += amount

    for c, title in enumerate(['CPF_Normalizado', 'Honorario_Valido_Recebido',
                               'Quantidade_Recebimentos', 'Sem_Honorario_Valido', 'Esperado_Sem_Recebimento'], 10):
        ps.cell(1, c, title)
    missing_count = 0
    missing_amount = Decimal(0)
    for r, p in enumerate(processes, 2):
        expected = (Decimal(str(p[6])) * Decimal(str(p[7]))).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
        formula(ps, f'I{r}', f'=ROUND(G{r}*H{r},2)', expected)
        document = cpf(p[1])
        absent = receipt_count[document] == 0
        missing_count += absent
        missing_amount += expected if absent else Decimal(0)
        helper_values = [document, float(receipts[document]), receipt_count[document],
                         int(absent), float(expected if absent else 0)]
        helper_formulas = [cpf_formula(f'B{r}'),
            f'=SUMIFS(BASE_LANCAMENTOS!$M$2:$M${n},BASE_LANCAMENTOS!$K$2:$K${n},J{r},BASE_LANCAMENTOS!$P$2:$P${n},1,BASE_LANCAMENTOS!$D$2:$D${n},"Honorários de Êxito")',
            f'=COUNTIFS(BASE_LANCAMENTOS!$K$2:$K${n},J{r},BASE_LANCAMENTOS!$P$2:$P${n},1,BASE_LANCAMENTOS!$D$2:$D${n},"Honorários de Êxito")',
            f'=--(L{r}=0)', f'=IF(M{r}=1,I{r},0)']
        for c, (expr, value) in enumerate(zip(helper_formulas, helper_values), 10):
            formula(ps, f'{openpyxl.utils.get_column_letter(c)}{r}', expr, value)

    ar = w['CONTAS_A_RECEBER']
    installments = [row for row in ar.iter_rows(min_row=2, max_col=8, values_only=True) if isinstance(row[4], datetime)]
    for c, title in enumerate(['Unidade_Padronizada', 'Valor_Numerico', 'Elegivel_3T26',
                               'Inadimplente_30_09', 'Valor_Base', 'Valor_Inadimplente'], 9):
        ar.cell(1, c, title)
    arrears = {u: [Decimal(0), Decimal(0)] for u in UNITS}
    for r, row in enumerate(installments, 2):
        title, client, document, branch, due, value, received, status = row
        canonical, amount = unit(branch), money(value)
        eligible = datetime(2026, 7, 1) <= due <= datetime(2026, 9, 30) and norm(status) not in ('cancelado', 'renegociado')
        unpaid = eligible and (received is None or received > datetime(2026, 9, 30))
        denominator, numerator = (amount if eligible else Decimal(0)), (amount if unpaid else Decimal(0))
        arrears[canonical][0] += denominator
        arrears[canonical][1] += numerator
        helper_values = [canonical, float(amount), int(eligible), int(unpaid), float(denominator), float(numerator)]
        helper_formulas = [unit_formula(f'D{r}'), money_formula(f'F{r}'),
            f'=--AND(ISNUMBER(E{r}),E{r}>=DATE(2026,7,1),E{r}<=DATE(2026,9,30),TRIM(H{r})<>"Cancelado",TRIM(H{r})<>"Renegociado")',
            f'=--AND(K{r}=1,OR(G{r}="",G{r}>DATE(2026,9,30)))',
            f'=IF(K{r}=1,J{r},0)', f'=IF(L{r}=1,J{r},0)']
        for c, (expr, value) in enumerate(zip(helper_formulas, helper_values), 9):
            formula(ar, f'{openpyxl.utils.get_column_letter(c)}{r}', expr, value)

    dre = w['RESULTADO_UNIDADE']
    n = ls.max_row
    results = {}
    aggregate = defaultdict(Decimal)
    for r, u in enumerate(UNITS, 4):
        revenue, cost, expense = (totals[u][g] for g in ('Receita', 'Custo', 'Despesa'))
        result = revenue - cost - expense
        results[u] = result
        for g in ('Receita', 'Custo', 'Despesa'):
            aggregate[g] += totals[u][g]
        for col, group, value in [('B', 'Receita', revenue), ('C', 'Custo', cost), ('E', 'Despesa', expense)]:
            formula(dre, f'{col}{r}', f'=SUMIFS(BASE_LANCAMENTOS!$M$2:$M${n},BASE_LANCAMENTOS!$L$2:$L${n},$A{r},BASE_LANCAMENTOS!$P$2:$P${n},1,BASE_LANCAMENTOS!$Q$2:$Q${n},"{group}")', value)
        formula(dre, f'D{r}', f'=IFERROR((B{r}-C{r})/B{r},0)', (revenue-cost)/revenue if revenue else 0)
        formula(dre, f'F{r}', f'=B{r}-C{r}-E{r}', result)
        formula(dre, f'G{r}', f'=IFERROR(F{r}/B{r},0)', result/revenue if revenue else 0)
    revenue, cost, expense = (aggregate[g] for g in ('Receita', 'Custo', 'Despesa'))
    result = revenue-cost-expense
    for col, value in [('B', revenue), ('C', cost), ('E', expense), ('F', result)]:
        formula(dre, f'{col}10', f'=SUM({col}4:{col}9)', value)
    formula(dre, 'D10', '=IFERROR((B10-C10)/B10,0)', (revenue-cost)/revenue)
    formula(dre, 'G10', '=IFERROR(F10/B10,0)', result/revenue)
    dre['D3'] = 'Margem de Contribuição %'

    aud = w.create_sheet('MEMORIA_CALCULO')
    notes = [
        ['Premissa / controle', 'Critério'],
        ['Fonte', Path(source).name],
        ['Período e regime', '01/07 a 30/09/2026; caixa dos lançamentos válidos. Não representa DRE por competência.'],
        ['Duplicação', 'CPF/CNPJ sem pontuação + data + categoria normalizada + valor em centavos. Primeira ocorrência na ordem original; deduplicação antes do filtro de status.'],
        ['Unidades', 'Ignorar espaços e caixa; Sede/Recife → Recife; Afogados da Ingazeira → Afogados; Santa Cruz do Capibaribe e Sta Cruz → Santa Cruz. Vazios → Não classificado. Sem rateio.'],
        ['Valores texto', 'Contar pelo tipo original, inclusive cancelados e cópias. Converter R$, separadores e espaços com Decimal e duas casas.'],
        ['Classificação', 'Custas, laudos/perícias e comissões são custos diretos. Demais saídas são despesas de suporte ou comerciais.'],
        ['Produtos e RPVs', 'Junção pelo CPF, sem exigir igualdade de valor. CPF único na base de processos (validado). Q12 conta processos sem nenhum honorário de êxito válido, não diferenças ou pagamentos parciais.'],
        ['Inadimplência', 'Vencimentos inclusivos jul–set; excluir cancelados e renegociados. Data de recebimento vazia ou posterior a 30/09 é inadimplente naquela data, mesmo recebida em outubro.'],
        ['Precisão', 'Valores em centavos; percentuais calculados sem arredondamento intermediário. Margens consolidadas ponderadas pela receita.'],
        ['Atualização', 'Auxiliares, DRE e respostas numéricas usam fórmulas. Alterações nas linhas existentes recalculam no Excel; ao inserir linhas ou novos produtos, ampliar intervalos ou executar o script. Recomendações textuais registram a análise inicial.'],
        ['Autoria', 'Declaração de autoria deixada em branco para leitura e preenchimento pessoal pelo candidato.'],
    ]
    for row in notes:
        aud.append(row)
    aud.append(['Controle', 'Quantidade', 'Valor R$'])
    aud.append(['Lançamentos originais', len(entries)])
    aud.append(['Cópias duplicadas', duplicate_count, float(duplicate_amount)])
    aud.append(['Cancelados (independente de duplicação)', cancelled_count])
    aud.append(['Valores originalmente texto', text_count])
    aud.append(['Lançamentos válidos', valid_count])
    aud.append(['RPVs sem nenhum honorário válido', missing_count, float(missing_amount)])
    for address, expr, value in [
        ('B14', f'=COUNTA(BASE_LANCAMENTOS!A2:A{n})', len(entries)),
        ('B15', f'=SUM(BASE_LANCAMENTOS!O2:O{n})', duplicate_count),
        ('C15', f'=SUMIF(BASE_LANCAMENTOS!O2:O{n},1,BASE_LANCAMENTOS!M2:M{n})', duplicate_amount),
        ('B16', f'=COUNTIF(BASE_LANCAMENTOS!I2:I{n},"Cancelado")', cancelled_count),
        ('B17', f'=SUM(BASE_LANCAMENTOS!N2:N{n})', text_count),
        ('B18', f'=SUM(BASE_LANCAMENTOS!P2:P{n})', valid_count),
        ('B19', f'=SUM(BASE_PROCESSOS!M2:M{pn})', missing_count),
        ('C19', f'=SUM(BASE_PROCESSOS!N2:N{pn})', missing_amount)]:
        formula(aud,address,expr,value)

    prod = w.create_sheet('RECEITA_PRODUTO')
    prod.append(['Produto', 'Honorários de êxito válidos'])
    sorted_products = sorted(product_totals.items(), key=lambda kv: kv[1], reverse=True)
    for r, (product, amount) in enumerate(sorted_products, 2):
        prod.cell(r, 1, product)
        formula(prod, f'B{r}', f'=SUMIFS(BASE_LANCAMENTOS!$M$2:$M${n},BASE_LANCAMENTOS!$R$2:$R${n},A{r},BASE_LANCAMENTOS!$P$2:$P${n},1,BASE_LANCAMENTOS!$D$2:$D${n},"Honorários de Êxito")', amount)
    debt = w.create_sheet('INADIMPLENCIA')
    debt.append(['Unidade', 'Parcelas elegíveis R$', 'Inadimplente 30/09 R$', 'Índice'])
    an = ar.max_row
    for r, u in enumerate(UNITS, 2):
        den, num = arrears[u]
        debt.cell(r, 1, u)
        for col, source_col, value in [('B', 'M', den), ('C', 'N', num)]:
            formula(debt, f'{col}{r}', f'=SUMIF(CONTAS_A_RECEBER!$I$2:$I${an},A{r},CONTAS_A_RECEBER!${source_col}$2:${source_col}${an})', value)
        formula(debt, f'D{r}', f'=IFERROR(C{r}/B{r},0)', num/den if den else 0)
    den = sum(v[0] for v in arrears.values())
    num = sum(v[1] for v in arrears.values())
    debt['A8'] = 'TOTAL'
    formula(debt, 'B8', '=SUM(B2:B7)', den)
    formula(debt, 'C8', '=SUM(C2:C7)', num)
    formula(debt, 'D8', '=IFERROR(C8/B8,0)', num/den)
    best = max(UNITS[:5], key=lambda u: results[u])
    worst = min(UNITS[:5], key=lambda u: results[u])
    most_debt = max(UNITS[:5], key=lambda u: arrears[u][1]/arrears[u][0])
    answers = w['RESPOSTAS']
    def answer(row, col, expression, value):
        formula(answers, f'{col}{row}', expression, value)
    answer(6, 'D', f'=SUM(BASE_LANCAMENTOS!O2:O{n})', duplicate_count)
    answer(6, 'F', f'=SUMIF(BASE_LANCAMENTOS!O2:O{n},1,BASE_LANCAMENTOS!M2:M{n})', duplicate_amount)
    answer(7, 'D', f'=COUNTIF(BASE_LANCAMENTOS!I2:I{n},"Cancelado")', cancelled_count)
    answer(8, 'D', f'=SUM(BASE_LANCAMENTOS!N2:N{n})', text_count)
    answers['D10'] = 'Custo é o gasto diretamente ligado à execução dos serviços jurídicos.\nDespesa sustenta a administração, a estrutura e a comercialização do escritório.'
    for row, col, expr, value in [(11,'D',"=RESULTADO_UNIDADE!C10",cost), (11,'F',"=RESULTADO_UNIDADE!E10",expense),
        (13,'D',"=RESULTADO_UNIDADE!B10",revenue), (13,'F',"=RESULTADO_UNIDADE!F10",result),
        (14,'D',"=RESULTADO_UNIDADE!D10",(revenue-cost)/revenue), (14,'F',"=RESULTADO_UNIDADE!G10",result/revenue),
        (15,'D','=INDEX(RESULTADO_UNIDADE!A4:A8,MATCH(MAX(RESULTADO_UNIDADE!F4:F8),RESULTADO_UNIDADE!F4:F8,0))',best),
        (15,'F','=INDEX(RESULTADO_UNIDADE!A4:A8,MATCH(MIN(RESULTADO_UNIDADE!F4:F8),RESULTADO_UNIDADE!F4:F8,0))',worst),
        (17,'D',f'=INDEX(RECEITA_PRODUTO!A2:A{prod.max_row},MATCH(MAX(RECEITA_PRODUTO!B2:B{prod.max_row}),RECEITA_PRODUTO!B2:B{prod.max_row},0))',sorted_products[0][0]),
        (17,'F',f'=MAX(RECEITA_PRODUTO!B2:B{prod.max_row})',sorted_products[0][1]),
        (19,'D',"=INADIMPLENCIA!D8",num/den),(19,'F',"=INADIMPLENCIA!C8",num),
        (20,'D','=INDEX(INADIMPLENCIA!A2:A6,MATCH(MAX(INADIMPLENCIA!D2:D6),INADIMPLENCIA!D2:D6,0))',most_debt),
        (20,'F','=MAX(INADIMPLENCIA!D2:D6)',arrears[most_debt][1]/arrears[most_debt][0]),
        (22,'D',f'=SUM(BASE_PROCESSOS!M2:M{ps.max_row})',missing_count),
        (22,'F',f'=SUM(BASE_PROCESSOS!N2:N{ps.max_row})',missing_amount)]:
        answer(row,col,expr,value)
    memories = {6:'BASE_LANCAMENTOS!K:S; chave, primeira linha, flag O e valor M. Script: solve, laço entries.',
        7:'BASE_LANCAMENTOS!I:I; contagem de todos os cancelados, inclusive eventuais cópias.',
        8:'BASE_LANCAMENTOS!N:N; ISTEXT sobre o valor original, antes de conversão.',
        11:'CLASSIFICACAO!C2:D13; RESULTADO_UNIDADE!C10 e E10; apenas P=1.',
        13:'RESULTADO_UNIDADE!B10 e F10; inclui Não classificado, sem rateio.',
        14:'RESULTADO_UNIDADE!D10 e G10; divisão pela receita consolidada.',
        15:'RESULTADO_UNIDADE!F4:F8; exclui Não classificado do ranking.',
        17:'RECEITA_PRODUTO!A:B; BASE_LANCAMENTOS!R:R; junção CPF único, somente êxito válido.',
        19:'INADIMPLENCIA!B8:D8; CONTAS_A_RECEBER!I:N. Corte em 30/09/2026.',
        20:'INADIMPLENCIA!D2:D6; ranking somente das cinco unidades oficiais.',
        22:'BASE_PROCESSOS!J:N; flag M = nenhum êxito válido pelo CPF; N = esperado nessas RPVs.'}
    for row, text in memories.items():
        answers.cell(row, 7, text)
    def brl(value):
        value = value.quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
        return f'R$ {value:,.2f}'.replace(',', '_').replace('.', ',').replace('_', '.')
    # Excluo recebimentos posteriores do saldo em aberto projetado.
    still_open = sum(money(r[5]) for r in installments if datetime(2026,7,1) <= r[4] <= datetime(2026,9,30)
                     and norm(r[7]) not in ('cancelado','renegociado') and r[6] is None)
    recovered_later = num-still_open
    reducible = sum(totals[u]['Despesa'] for u in UNITS)
    answers['D24'] = (f'1) Conciliar e cobrar RPVs sem êxito: potencial bruto de {brl(missing_amount)} ({missing_count} processos), condicionado à confirmação do direito e recebimento.\n'
        f'2) Priorizar a carteira vencida: {brl(num)} em 30/09; {brl(recovered_later)} já recebidos depois do corte. Meta de recuperar 25% dos {brl(still_open)} sem data de recebimento: {brl(still_open*Decimal(".25"))} em caixa.\n'
        f'3) Revisar contratos e gastos de suporte: meta de reduzir 5% das despesas trimestrais de {brl(reducible)}, elevando o resultado em {brl(reducible*Decimal(".05"))}, mantidas receita e operação. Metas são cenários, não ganhos garantidos; não somar carteiras sem conciliar.')
    answers['D26'] = 'resolver_desafio.py (anexo); código também na aba PYTHON, a partir de B8.'
    for row in (6,7,8,22):
        answers.cell(row,4).number_format='0'
    for row,col in [(6,6),(11,4),(11,6),(13,4),(13,6),(17,6),(19,6),(22,6)]:
        answers.cell(row,col).number_format='#,##0.00'
    for row,col in [(14,4),(14,6),(19,4),(20,6)]:
        answers.cell(row,col).number_format='0.0%'
    formula(answers, 'H27', '=SUM(H6:H24)', sum(answers.cell(r,8).value or 0 for r in range(6,25)))
    for r in range(4,11):
        for col in ('B','C','E','F'):
            dre[f'{col}{r}'].number_format='#,##0.00'
        for col in ('D','G'):
            dre[f'{col}{r}'].number_format='0.0%'
    for r in range(2,9):
        debt[f'D{r}'].number_format='0.0%'
        for col in ('B','C'):
            debt[f'{col}{r}'].number_format='#,##0.00'
    for r in range(2,prod.max_row+1):
        prod[f'B{r}'].number_format='#,##0.00'
    for sheet, cols in [(ls,[13]),(ps,[11,14]),(ar,[10,13,14])]:
        for r in range(2,sheet.max_row+1):
            for c in cols:
                sheet.cell(r,c).number_format='#,##0.00'
    for sheet in [ls,ps,ar,prod,debt,aud]:
        sheet.freeze_panes='A2'
        sheet.auto_filter.ref=sheet.dimensions
        for cell in sheet[1]:
            cell.fill=PatternFill('solid',fgColor='17365D')
            cell.font=Font(color='FFFFFF',bold=True)
        for row in sheet:
            for cell in row:
                if cell.column >= (11 if sheet==ls else 10 if sheet==ps else 9 if sheet==ar else 1):
                    cell.alignment=Alignment(vertical='top',wrap_text=True)
    for sheet,start,end in [(ls,11,19),(ps,10,14),(ar,9,14)]:
        for c in range(start,end+1):
            sheet.column_dimensions[openpyxl.utils.get_column_letter(c)].width=25
    aud.column_dimensions['A'].width=43
    aud.column_dimensions['B'].width=115
    aud.column_dimensions['C'].width=22
    for r in range(2,13):
        aud.row_dimensions[r].height=48
    prod.column_dimensions['A'].width=52
    prod.column_dimensions['B'].width=28
    for col in ('A','B','C','D'):
        debt.column_dimensions[col].width=29
    answers.row_dimensions[24].height=180
    answers.row_dimensions[10].height=55
    answers['D24'].alignment=Alignment(wrap_text=True,vertical='top')
    answers['D10'].alignment=Alignment(wrap_text=True,vertical='top')
    for r in memories:
        answers.cell(r,7).alignment=Alignment(wrap_text=True,vertical='top')
        answers.row_dimensions[r].height=max(answers.row_dimensions[r].height or 15,55)
    formatar_apresentacao(w)
    code = Path(__file__).read_text(encoding='utf-8')
    for r,line in enumerate(code.splitlines(),8):
        w['PYTHON'].cell(r,2,line)
    w['PYTHON'].column_dimensions['B'].width=135
    w.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True)
    w.save(destination)
    # Gravo no cache OOXML os resultados calculados, preservando as fórmulas.
    namespace='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    tag=lambda s: '{'+namespace+'}'+s
    with zipfile.ZipFile(destination) as archive:
        files={item.filename: archive.read(item.filename) for item in archive.infolist()}
    for index,sheet in enumerate(w,1):
        path=f'xl/worksheets/sheet{index}.xml'
        root=ET.fromstring(files[path])
        for cell in root.iter(tag('c')):
            key=(sheet.title,cell.attrib['r'])
            if key not in cache:
                continue
            value=cache[key]
            v=cell.find(tag('v'))
            if v is None:
                v=ET.SubElement(cell,tag('v'))
            if isinstance(value,str):
                cell.set('t','str')
                v.text=value
            else:
                cell.attrib.pop('t',None)
                v.text=str(float(value)) if isinstance(value,Decimal) else str(value)
        files[path]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
    with zipfile.ZipFile(destination,'w',zipfile.ZIP_DEFLATED) as archive:
        for path,data in files.items():
            archive.writestr(path,data)
    # Valido a preservação das bases, os caches e a reconciliação dos totais.
    check=openpyxl.load_workbook(destination,data_only=True)
    original=openpyxl.load_workbook(source,data_only=False)
    saved=openpyxl.load_workbook(destination,data_only=False)
    for name,columns in [('BASE_LANCAMENTOS',10),('BASE_PROCESSOS',8),('CONTAS_A_RECEBER',8)]:
        for row in original[name].iter_rows(max_col=columns):
            for cell in row:
                assert saved[name][cell.coordinate].value==cell.value,(name,cell.coordinate)
    for (name,address),expected in cache.items():
        actual=check[name][address].value
        assert (actual or '')==expected if isinstance(expected,str) else abs(float(actual)-float(expected))<1e-7,(name,address)
    assert sum(product_totals.values())==sum(receipts.values())
    assert sum(results.values())==result
    assert den>0 and 0<=num<=den
    assert check['RESPOSTAS']['D33'].value is None
    print('Arquivo gerado e validado:',destination)
    for r in [6,7,8,11,13,14,15,17,19,20,22]:
        print(check['RESPOSTAS'].cell(r,1).value,check['RESPOSTAS'].cell(r,4).value,check['RESPOSTAS'].cell(r,6).value)
    return check


if __name__=='__main__':
    solve(sys.argv[1] if len(sys.argv)>1 else 'Desafio_Controller_Italo (2).xlsx',
          sys.argv[2] if len(sys.argv)>2 else 'Desafio_Italo_Andre_Mendes_da_Silva_Ramos_Completo.xlsx')
