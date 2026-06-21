#!/usr/bin/env python3
"""Unifica planilhas NR-13 da raiz do repositório em um XLSX auditável."""
from __future__ import annotations
from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from datetime import datetime, timezone, timedelta
import re, unicodedata, html, xml.etree.ElementTree as ET, hashlib

OUT = 'Engetap_NR13_Unificado_Profissional.xlsx'
SHEETS = ['Painel','Base_Limpa','Base_Bruta','Vasos','Tanques','Tubulações','Pendencias','Dicionario_Campos','Log_Importacao']
AUDIT = ['Fonte_Arquivo','Fonte_Aba','Linha_Origem','Unidade_Arquivo','Area_Arquivo','OS_Normalizada','Tipo_Normalizado','TAG_Normalizada','Data_Coleta_Formatada','Flag_Placeholder_76_77','Colunas_Com_Placeholder','Flag_Dados_Incompletos','Flag_Duplicidade_Potencial','Observacao_Auditoria']
BASE_ORIGIN = ['Fonte_Arquivo','Fonte_Aba','Linha_Origem']

def repo_root(): return Path(__file__).resolve().parents[1]
def norm(s):
    s='' if s is None else str(s)
    s=unicodedata.normalize('NFKD',s).encode('ascii','ignore').decode().lower()
    s=re.sub(r'[\r\n]+',' ',s); s=re.sub(r'[^a-z0-9]+','_',s); return re.sub(r'_+','_',s).strip('_')
def clean(v):
    if v is None: return ''
    if isinstance(v,float) and v.is_integer(): return str(int(v))
    return str(v).strip()
def colnum(ref):
    n=0
    for ch in re.sub(r'\d','',ref): n=n*26+ord(ch.upper())-64
    return n

def find_inputs(root):
    return sorted(p for p in root.glob('*.xlsx') if not p.name.startswith('~$') and p.name != OUT and 'NR13' in norm(p.name).upper().replace('_',''))

def read_xlsx(path):
    ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main','r':'http://schemas.openxmlformats.org/officeDocument/2006/relationships'}
    with ZipFile(path) as z:
        wb=ET.fromstring(z.read('xl/workbook.xml')); rels=ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
        rid=None
        for s in wb.findall('m:sheets/m:sheet',ns):
            if s.attrib.get('name')=='NR-13': rid=s.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']
        if not rid: return None, 'Aba NR-13 não encontrada'
        target=None
        for r in rels:
            if r.attrib.get('Id')==rid: target=r.attrib['Target']
        sheet_path='xl/'+target.lstrip('/') if not target.startswith('xl/') else target
        shared=[]
        if 'xl/sharedStrings.xml' in z.namelist():
            ss=ET.fromstring(z.read('xl/sharedStrings.xml'))
            for si in ss.findall('m:si',ns): shared.append(''.join(t.text or '' for t in si.findall('.//m:t',ns)))
        root=ET.fromstring(z.read(sheet_path)); rows=[]
        for row in root.findall('.//m:sheetData/m:row',ns):
            vals={}; maxc=0
            for c in row.findall('m:c',ns):
                idx=colnum(c.attrib.get('r','A1')); maxc=max(maxc,idx); t=c.attrib.get('t')
                v=c.find('m:v',ns); isel=c.find('m:is',ns)
                val=''
                if t=='s' and v is not None: val=shared[int(v.text)] if v.text and int(v.text)<len(shared) else ''
                elif t=='inlineStr' and isel is not None: val=''.join(x.text or '' for x in isel.findall('.//m:t',ns))
                elif v is not None: val=v.text or ''
                vals[idx]=val
            if vals: rows.append([vals.get(i,'') for i in range(1,maxc+1)])
        if not rows: return [], None
        return rows, None

def best(row, headers, names):
    want=[norm(x) for x in names]
    for h in headers:
        nh=norm(h)
        if any(w==nh or w in nh for w in want): return row.get(h,'')
    return ''
def normalize_os(v): return re.sub(r'-+','-',re.sub(r'[_\s]+','-',clean(v).upper())).strip('-')
def unidade(fname,row,headers):
    text=fname+' '+best(row,headers,['OS'])
    m=re.search(r'\b(FCC|PHF|ALL)\b',text.upper()); return m.group(1) if m else 'OUTRA'
def area(fname,row,headers):
    t=norm(fname+' '+best(row,headers,['Setor','Localizacao','Área']))
    pairs=[('TUBULAÇÕES','tubul'),('ÁREA DE TANQUES','tanque|tqs'),('PET','pet'),('EMBALAGEM','embal'),('UTILIDADES','util'),('SETORES VARIADOS','setores_variados'),('NÃO ENQUADRADOS','nao_enquadr')]
    for label,pat in pairs:
        if re.search(pat,t): return label
    return clean(best(row,headers,['Setor','Área'])) or 'OUTRA'
def tipo(fname,row,headers):
    t=norm(fname+' '+best(row,headers,['Tipo','Equipamento','Descrição','TAG','Setor']))
    if 'tubul' in t: return 'tubulacao'
    if 'tanque' in t or 'tqs' in t: return 'tanque'
    if any(x in t for x in ['vaso','utilidade','pet','embalagem']): return 'vaso'
    if best(row,headers,['PMTA','PSV','P_Trabalho']): return 'vaso'
    return 'indefinido'
def fmt_date(v):
    s=clean(v)
    if not s: return ''
    try:
        if re.fullmatch(r'\d+(\.\d+)?',s):
            d=datetime(1899,12,30)+timedelta(days=float(s)); return d.strftime('%d/%m/%Y %H:%M')
    except Exception: pass
    return s

def load_all(root):
    headers=[]; bruta=[]; logs=[]; ignored=[]
    for p in find_inputs(root):
        rows,err=read_xlsx(p)
        if err: ignored.append((p.name,err)); continue
        if not rows: ignored.append((p.name,'Sem linhas')); continue
        hdr=[clean(x) or f'Coluna_{i+1}' for i,x in enumerate(rows[0])]
        for h in hdr:
            if h not in headers: headers.append(h)
        count=0; ph_rows=0; ph_total=0
        for rn,vals in enumerate(rows[1:], start=2):
            if not any(clean(v) for v in vals): continue
            row={h: clean(vals[i]) if i < len(vals) else '' for i,h in enumerate(hdr)}
            row.update({'Fonte_Arquivo':p.name,'Fonte_Aba':'NR-13','Linha_Origem':rn}); bruta.append(row); count+=1
            cols=[h for h in hdr if clean(row.get(h)) in ('76','77')]; ph_total+=len(cols); ph_rows+=1 if cols else 0
        logs.append({'Arquivo':p.name,'Aba':'NR-13','Linhas_Importadas':count,'Colunas_Encontradas':len(hdr),'Cabecalhos_Encontrados':', '.join(hdr),'Cabecalhos_Padrao':'Sim' if len(hdr)>=10 else 'Não','Placeholders_76_77':ph_total,'Linhas_Com_Placeholders':ph_rows,'Avisos':'','Data_Execucao':datetime.now().strftime('%d/%m/%Y %H:%M:%S')})
    return headers, bruta, logs, ignored

def build_clean(headers, bruta):
    out=[]
    for r in bruta:
        osn=normalize_os(best(r,headers,['OS','Ordem Serviço','Ordem de Serviço']))
        tpn=tipo(r['Fonte_Arquivo'],r,headers); tag=clean(best(r,headers,['TAG','Tag Equipamento','Identificação'])).upper()
        ph=[h for h in headers if clean(r.get(h)) in ('76','77')]
        incomplete='Sim' if not tag or tpn=='indefinido' else 'Não'
        obs=[]
        if tpn=='indefinido': obs.append('Tipo não inferido com segurança')
        if ph: obs.append('Contém placeholders suspeitos 76/77')
        nr={k:'' for k in AUDIT}; nr.update({'Fonte_Arquivo':r['Fonte_Arquivo'],'Fonte_Aba':'NR-13','Linha_Origem':r['Linha_Origem'],'Unidade_Arquivo':unidade(r['Fonte_Arquivo'],r,headers),'Area_Arquivo':area(r['Fonte_Arquivo'],r,headers),'OS_Normalizada':osn,'Tipo_Normalizado':tpn,'TAG_Normalizada':tag,'Data_Coleta_Formatada':fmt_date(best(r,headers,['Data_Coleta','Data','Data coleta'])),'Flag_Placeholder_76_77':'Sim' if ph else 'Não','Colunas_Com_Placeholder':', '.join(ph),'Flag_Dados_Incompletos':incomplete,'Flag_Duplicidade_Potencial':'Não','Observacao_Auditoria':'; '.join(obs)})
        for h in headers: nr[h]=r.get(h,'')
        out.append(nr)
    keys={}
    for r in out:
        for key in [(r['OS_Normalizada'],r['TAG_Normalizada']),(best(r,headers,['Cliente']),r['TAG_Normalizada'],r['Tipo_Normalizado'])]:
            if all(key): keys.setdefault(key,[]).append(r)
    for rows in keys.values():
        if len(rows)>1:
            for r in rows: r['Flag_Duplicidade_Potencial']='Sim'
    return out

def rows_for(headers, rows): return [headers]+[[r.get(h,'') for h in headers] for r in rows]
def count_by(rows,key):
    d={}
    for r in rows: d[r.get(key,'') or '(vazio)']=d.get(r.get(key,'') or '(vazio)',0)+1
    return sorted(d.items(), key=lambda x:(-x[1],x[0]))
def pendente(r, headers):
    text=norm(' '.join(clean(r.get(h,'')) for h in headers if any(x in norm(h) for x in ['pendenc','status_doc','document'])))
    no_drive=not best(r,headers,['Drive_URL','Drive','Link'])
    fotos=best(r,headers,['Qtd_Fotos','Fotos']); low=(not fotos) or fotos in ('0','1')
    return r['Flag_Placeholder_76_77']=='Sim' or r['Flag_Dados_Incompletos']=='Sim' or r['Flag_Duplicidade_Potencial']=='Sim' or no_drive or low or any(x in text for x in ['pendente','ausente','parcial','nao_localizada','nao_localizado'])

def xlsx_escape(s): return html.escape(clean(s), quote=True)
def write_xlsx(path, sheets):
    def col_letter(n):
        s=''
        while n: n,r=divmod(n-1,26); s=chr(65+r)+s
        return s
    with ZipFile(path,'w',ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml','<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'+''.join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' for i in range(1,len(sheets)+1))+'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
        z.writestr('_rels/.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr('xl/_rels/workbook.xml.rels','<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'+''.join(f'<Relationship Id="rId{i}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{i}.xml"/>' for i in range(1,len(sheets)+1))+'<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        z.writestr('xl/workbook.xml','<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'+''.join(f'<sheet name="{xlsx_escape(name)}" sheetId="{i}" r:id="rId{i}"/>' for i,name in enumerate(sheets,1))+'</sheets></workbook>')
        z.writestr('xl/styles.xml','<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font/><font><b/><color rgb="FFFFFFFF"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf fontId="1" fillId="2" applyFont="1" applyFill="1"/></cellXfs></styleSheet>')
        for idx,(name,data) in enumerate(sheets.items(),1):
            maxc=max((len(r) for r in data), default=1); rows=[]
            for ri,row in enumerate(data,1):
                cells=[]
                for ci,val in enumerate(row,1):
                    style=' s="1"' if ri==1 else ''
                    cells.append(f'<c r="{col_letter(ci)}{ri}" t="inlineStr"{style}><is><t>{xlsx_escape(val)}</t></is></c>')
                rows.append(f'<row r="{ri}">'+''.join(cells)+'</row>')
            z.writestr(f'xl/worksheets/sheet{idx}.xml',f'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><dimension ref="A1:{col_letter(maxc)}{max(len(data),1)}"/><sheetData>'+''.join(rows)+'</sheetData><autoFilter ref="A1:{col_letter(maxc)}{max(len(data),1)}"/></worksheet>')

def main():
    root=repo_root(); (root/'outputs').mkdir(exist_ok=True); (root/'reports').mkdir(exist_ok=True)
    before={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in root.glob('*.xlsx') if p.name!=OUT}
    headers, bruta, logs, ignored=load_all(root); limpa=build_clean(headers,bruta)
    origin_headers=BASE_ORIGIN+headers; clean_headers=AUDIT+headers
    painel=[['Indicador','Valor'],['Total de arquivos processados',len(logs)],['Total de registros importados',len(bruta)],['Registros com pendências',sum(pendente(r,headers) for r in limpa)],['Placeholders 76/77',sum(r['Flag_Placeholder_76_77']=='Sim' for r in limpa)],['Duplicidades potenciais',sum(r['Flag_Duplicidade_Potencial']=='Sim' for r in limpa)],[],['Por unidade','Total']]+count_by(limpa,'Unidade_Arquivo')+[[],['Por tipo','Total']]+count_by(limpa,'Tipo_Normalizado')+[[],['Por área','Total']]+count_by(limpa,'Area_Arquivo')
    def view(tp): return [r for r in limpa if r['Tipo_Normalizado']==tp]
    pend=[r for r in limpa if pendente(r,headers)]
    dic=[['Campo','Descrição','Tipo_Esperado','Exemplo','Aplicável_a','Observações'],['Fonte_Arquivo','Arquivo de origem','texto','Engetap_NR13...xlsx','todos','Rastreabilidade'],['Linha_Origem','Linha original na aba NR-13','inteiro','2','todos','Rastreabilidade'],['Tipo_Normalizado','Classificação inferida','texto','vaso','todos','vaso/tanque/tubulacao/indefinido'],['Flag_Placeholder_76_77','Marca valores suspeitos 76/77','Sim/Não','Sim','todos','Não apaga dados brutos'],['Flag_Duplicidade_Potencial','Repetição por chaves técnicas','Sim/Não','Sim','todos','Não exclui registros']]
    logh=['Arquivo','Aba','Linhas_Importadas','Colunas_Encontradas','Cabecalhos_Encontrados','Cabecalhos_Padrao','Placeholders_76_77','Linhas_Com_Placeholders','Avisos','Data_Execucao']
    data={'Painel':painel,'Base_Limpa':rows_for(clean_headers,limpa),'Base_Bruta':rows_for(origin_headers,bruta),'Vasos':rows_for(clean_headers,view('vaso')),'Tanques':rows_for(clean_headers,view('tanque')),'Tubulações':rows_for(clean_headers,view('tubulacao')),'Pendencias':rows_for(clean_headers,pend),'Dicionario_Campos':dic,'Log_Importacao':rows_for(logh,logs)}
    write_xlsx(root/'outputs'/OUT,data)
    after={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in root.glob('*.xlsx') if p.name!=OUT}
    unchanged=before==after
    report=['# Relatório de Unificação NR-13','',f'Gerado em: {datetime.now().strftime("%d/%m/%Y %H:%M:%S")}.','', '## Objetivo','Consolidar as planilhas NR-13 da raiz do repositório em uma base única, auditável e rerodável.','', '## Arquivos processados']
    report += [f'- {l["Arquivo"]}: {l["Linhas_Importadas"]} registros, {l["Colunas_Encontradas"]} colunas.' for l in logs]
    report += ['', '## Arquivos ignorados'] + ([f'- {a}: {m}' for a,m in ignored] or ['- Nenhum.'])
    report += ['', '## Totais', f'- Planilhas processadas: {len(logs)}', f'- Registros consolidados: {len(bruta)}', f'- Registros com placeholders 76/77: {sum(r["Flag_Placeholder_76_77"]=="Sim" for r in limpa)}', f'- Registros com pendências: {len(pend)}', f'- Duplicidades potenciais: {sum(r["Flag_Duplicidade_Potencial"]=="Sim" for r in limpa)}', '', '## Quantidade por tipo'] + [f'- {k}: {v}' for k,v in count_by(limpa,'Tipo_Normalizado')] + ['', '## Quantidade por unidade'] + [f'- {k}: {v}' for k,v in count_by(limpa,'Unidade_Arquivo')] + ['', '## Regras de normalização aplicadas','- Cabeçalhos normalizados internamente para busca flexível, preservando nomes originais nas abas de dados.','- OS, unidade, área, tipo, TAG e datas recebem colunas auxiliares sem alterar a base bruta.','- Valores 76/77 são preservados na Base_Bruta e marcados na Base_Limpa.','- Duplicidades potenciais são marcadas, nunca excluídas.','', '## Limitações conhecidas','- Inferências dependem dos nomes de arquivos, cabeçalhos e conteúdo disponível; campos ausentes não são inventados.','- O ambiente atual não tinha openpyxl instalado via rede, portanto o gerador usa OOXML padrão com biblioteca nativa do Python.','', '## Como rerodar','```bash','python scripts/unificar_nr13.py','```','',f'Arquivos originais preservados: {"Sim" if unchanged else "Não"}.']
    (root/'reports'/'RELATORIO_UNIFICACAO_NR13.md').write_text('\n'.join(report)+'\n', encoding='utf-8')
    assert len(bruta)==len(limpa)
    assert all(r['Fonte_Arquivo'] and r['Linha_Origem'] for r in bruta)
    print(f'Planilhas processadas: {len(logs)}')
    print(f'Registros consolidados: {len(bruta)}')
    print(f'Pendências: {len(pend)}')
    print(f'Placeholders 76/77: {sum(r["Flag_Placeholder_76_77"]=="Sim" for r in limpa)}')
    print(f'Originais preservados: {unchanged}')
if __name__=='__main__': main()
