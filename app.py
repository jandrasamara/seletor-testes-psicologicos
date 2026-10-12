import re
import unicodedata
from pathlib import Path
import pandas as pd
import streamlit as st

st.set_page_config(page_title='Seletor de instrumentos psicológicos', page_icon='🧠', layout='wide')
BASE = Path(__file__).with_name('Banco_SATEPSI_universal.xlsx')
SATEPSI = 'https://satepsi.cfp.org.br/'

def clean(x):
    return '' if pd.isna(x) else str(x).strip()

def key(x):
    return ' '.join(''.join(c for c in unicodedata.normalize('NFKD', clean(x).casefold()) if not unicodedata.combining(c)).split())

def get_data(file):
    book = pd.ExcelFile(file)
    required = ['Testes', 'Construtos', 'Modalidades', 'Formatos', 'Regras_Validadas']
    if all(x in book.sheet_names for x in required):
        result = {n: pd.read_excel(book, sheet_name=n).fillna('') for n in required}
    else:
        # Também aceita um banco simples com uma linha por instrumento.
        raw = pd.read_excel(book, sheet_name=0).fillna('')
        aliases = {'Nome_Teste':['Nome_Teste','Nome do Teste'], 'Construtos':['Construtos','Construto'],
                   'Idade_Minima':['Idade_Minima'], 'Idade_Maxima':['Idade_Maxima'],
                   'Escolaridade_Minima':['Escolaridade_Minima'], 'Forma_Aplicacao':['Forma_Aplicacao','Modalidade'],
                   'Formato':['Formato'], 'Observacoes_Original':['Observações','Observacoes','Observacoes_Original']}
        mapped = {}
        for target, choices in aliases.items():
            found = next((c for c in choices if c in raw.columns), None)
            mapped[target] = raw[found] if found else pd.Series(['']*len(raw))
        if not raw.columns.isin(aliases['Nome_Teste']).any() or not raw.columns.isin(aliases['Construtos']).any():
            raise ValueError('A planilha simples precisa das colunas Nome_Teste e Construtos.')
        tests = pd.DataFrame(mapped)
        tests = tests[tests['Nome_Teste'].map(lambda v: bool(clean(v)))].reset_index(drop=True)
        tests['ID_Teste'] = [f'T{i:04d}' for i in range(1,len(tests)+1)]
        for col in ['Situacao_SATEPSI','URL_Fonte','Data_Conferencia','Status_Revisao']:
            tests[col] = raw.loc[tests.index, col].reset_index(drop=True) if col in raw.columns else ''
        cr=[]; md=[]; fm=[]
        for _, r in tests.iterrows():
            for value in clean(r['Construtos']).split(';'):
                if value.strip(): cr.append({'ID_Teste':r['ID_Teste'],'Construto_Padronizado':value.strip()})
            for value in clean(r['Forma_Aplicacao']).split(';'):
                if value.strip(): md.append({'ID_Teste':r['ID_Teste'],'Modalidade':value.strip()})
            for value in clean(r['Formato']).split(';'):
                if value.strip(): fm.append({'ID_Teste':r['ID_Teste'],'Formato':value.strip()})
        result = {'Testes':tests,'Construtos':pd.DataFrame(cr,columns=['ID_Teste','Construto_Padronizado']),
                  'Modalidades':pd.DataFrame(md,columns=['ID_Teste','Modalidade']),
                  'Formatos':pd.DataFrame(fm,columns=['ID_Teste','Formato']),
                  'Regras_Validadas':pd.DataFrame(columns=['ID_Teste','Modalidade','Formato'])}
    if 'Dicionario_Construtos' in book.sheet_names:
        result['Dicionario_Construtos'] = pd.read_excel(book,sheet_name='Dicionario_Construtos').fillna('')
    else:
        result['Dicionario_Construtos'] = pd.DataFrame(columns=['Construto','Descricao','Contextos_de_Uso'])
    return result

# As explicações agora são fornecidas pelo banco Excel, nunca fixadas no código.
INFO = {}

SCHOOL = ['Não alfabetizado', 'Fundamental incompleto', 'Fundamental completo', 'Médio incompleto', 'Médio completo', 'Superior incompleto', 'Superior completo', 'Pós-graduação']

def min_school(value):
    """Return minimum schooling rank; blank/unrestricted = None; unknown = 'ambigua'."""
    k = key(value)
    if not k or any(z in k for z in ('independente', 'qualquer escolaridade', 'todos os niveis', 'sem exigencia', 'nao exige escolaridade', 'sem restricao')):
        return None
    if 'nao alfabetiz' in k or 'analfabet' in k: return 0
    if 'pos' in k and ('gradu' in k or 'doutor' in k): return 7
    if 'superior' in k or 'graduacao' in k or 'universitari' in k:
        return 5 if 'incompleto' in k else 6 if 'completo' in k else 5
    if 'medio' in k or 'segundo grau' in k:
        return 3 if 'incompleto' in k else 4 if 'completo' in k else 3
    if 'fundamental' in k or 'primeiro grau' in k:
        return 1 if 'incompleto' in k else 2 if 'completo' in k else 1
    if 'alfabetiz' in k: return 1
    return 'ambigua'

st.title('🧠 Apoio à seleção de testes psicológicos')
st.caption('Ferramenta acadêmica de apoio à triagem de instrumentos. Não realiza avaliação psicológica nem substitui o psicólogo.')
with st.sidebar:
    st.header('Base de instrumentos')
    st.markdown('**Modelos para preencher**')
    st.download_button('⬇️ Baixar modelo simples', data=(Path(__file__).parent / 'Modelo_Simples_Instrumentos.xlsx').read_bytes(), file_name='Modelo_Simples_Instrumentos.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', width='stretch')
    st.download_button('⬇️ Baixar modelo completo', data=(Path(__file__).parent / 'Modelo_Completo_Instrumentos.xlsx').read_bytes(), file_name='Modelo_Completo_Instrumentos.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', width='stretch')
    st.divider()
    st.markdown('**Banco em uso nesta aba**')
    uploaded = st.file_uploader('Anexar planilha preenchida (.xlsx)', type='xlsx', key='banco_da_aba')
    st.caption('Cada aba do navegador mantém sua própria sessão e seu banco importado. O arquivo não é salvo permanentemente no servidor.')
    st.caption('Use a aba Dicionario_Construtos para incluir as explicações dos construtos.')
    st.link_button('Consultar SATEPSI oficial', SATEPSI)
    st.info('A planilha não é uma conexão em tempo real com o SATEPSI. Confira a situação e as condições de uso no site oficial e no manual.')
try:
    db = get_data(uploaded if uploaded is not None else BASE)
except Exception as e:
    st.error(f'Erro ao abrir banco: {e}. Verifique se o arquivo Excel está na mesma pasta do app.py.')
    st.stop()
tests, cons, mods, forms, rules = (db[n] for n in ['Testes','Construtos','Modalidades','Formatos','Regras_Validadas'])
tests = tests[tests['Nome_Teste'].map(lambda v: bool(clean(v)))].copy()
st.sidebar.success(f'Banco ativo nesta aba: {uploaded.name if uploaded is not None else BASE.name} ({len(tests)} instrumentos)')
if uploaded is not None:
    st.caption(f'📂 Banco ativo nesta aba: **{uploaded.name}** — {len(tests)} instrumentos. Outras abas do navegador podem carregar outros arquivos.')
constructs = sorted({clean(v) for v in cons['Construto_Padronizado'] if clean(v)}, key=key)
dictionary = {}
for _, record in db['Dicionario_Construtos'].iterrows():
    if clean(record.get('Construto')):
        dictionary[key(record['Construto'])] = (clean(record.get('Descricao')), clean(record.get('Contextos_de_Uso')), clean(record.get('O_que_pode_mensurar')), clean(record.get('Cuidados_e_limites')))


with st.form('questionario'):
    st.subheader('1. Perfil da pessoa avaliada')
    a,b = st.columns(2)
    with a: age = st.number_input('Idade (anos)', min_value=0, max_value=120, value=25)
    with b: school = st.selectbox('Escolaridade', SCHOOL)
    st.subheader('2. Condições de aplicação')
    a,b = st.columns(2)
    with a: mode = st.multiselect('Modalidade (selecione uma ou mais)', ['Individual','Coletiva'], default=['Individual'])
    with b: fmt = st.multiselect('Formato (selecione um ou mais)', ['Não informatizada','Informatizada presencial','Remota'], default=['Não informatizada'])
    st.subheader('3. Construtos que deseja avaliar')
    wanted = st.multiselect('Selecione um ou mais construtos', constructs)
    extra = st.radio('Aceita testes que também avaliam outros construtos?', ['Sim','Não'], horizontal=True)
    submit = st.form_submit_button('Encontrar testes', type='primary', width='stretch')

with st.expander('📖 O que cada construto avalia? Veja as explicações e contextos de uso'):
    st.caption('Textos fornecidos pelo responsável pela planilha. A definição operacional e as indicações de cada instrumento devem ser consultadas em seu manual. Contextos de uso não constituem autorização normativa.')
    for c in constructs:
        desc, use, measures, limits = dictionary.get(key(c), ('Definição ainda não cadastrada pelo responsável pelo banco.', 'Contexto de uso ainda não cadastrado.', '', ''))
        st.markdown(f'**{c}** — {desc}  \n**O que pode mensurar:** {measures or "Conforme o manual do instrumento."}  \n*Contextos possíveis:* {use}  \n*Cuidados:* {limits or "Conferir normas e manual."}')
    st.info('Em trânsito, porte de arma, investigação de TDAH ou TEA e outras finalidades regulamentadas, a escolha depende das normas aplicáveis, da demanda e da avaliação profissional. Nenhum teste isolado estabelece diagnóstico de TDAH ou TEA.')

# Três estados para cada filtro: atende / não consta restrição / incompatível.
# O campo em branco é livre de restrição cadastrada, não autorização oficial.
def evaluate(t):
    tid=t['ID_Teste']; ok=[]; notes=[]; bad=[]
    cc=cons.loc[cons.ID_Teste==tid, 'Construto_Padronizado'].tolist()
    have={key(v) for v in cc}; need={key(v) for v in wanted}
    if not need.issubset(have):
        bad.append('Não contempla todos os construtos escolhidos')
    else:
        ok.append('Contempla os construtos escolhidos')
    if extra=='Não' and have-need:
        bad.append('Avalia outros construtos e você optou por não aceitá-los')
    for name,col,condition in [('idade mínima','Idade_Minima',lambda x:age>=x),('idade máxima','Idade_Maxima',lambda x:age<=x)]:
        v=pd.to_numeric(t[col],errors='coerce')
        if pd.isna(v):
            notes.append(f'Sem restrição de {name} cadastrada; conferir população-alvo e normas no manual')
        elif condition(v):
            ok.append(f'{name.capitalize()} compatível ({int(v)} anos)')
        else:
            bad.append(f'{name.capitalize()} incompatível (limite: {int(v)} anos)')
    raw_school=clean(t['Escolaridade_Minima'])
    ms=min_school(raw_school)
    if ms is None:
        ok.append('Escolaridade: sem exigência mínima cadastrada (filtro atendido)')
        # Sem restrição de escolaridade na base: não penaliza compatibilidade
    elif ms=='ambigua':
        notes.append(f'Escolaridade mínima não interpretada automaticamente: {raw_school}')
    elif SCHOOL.index(school)<ms:
        bad.append(f'Escolaridade informada ({school}) abaixo da mínima cadastrada ({raw_school})')
    else:
        ok.append(f'Escolaridade compatível com mínimo cadastrado: {raw_school}')
    mm={key(v) for v in mods.loc[mods.ID_Teste==tid,'Modalidade'] if clean(v)}
    ff={key(v) for v in forms.loc[forms.ID_Teste==tid,'Formato'] if clean(v)}
    selected_modes={key(v) for v in mode}
    selected_formats={key(v) for v in fmt}
    matched_modes=selected_modes & mm
    matched_formats=selected_formats & ff
    if not mm:
        ok.append('Sem restrição de modalidade cadastrada')
        notes.append('Modalidade não documentada na planilha: conferir manual')
    elif not matched_modes:
        bad.append('Nenhuma das modalidades solicitadas consta nas opções cadastradas')
    else:
        ok.append('Modalidade(s) compatível(is): '+', '.join(v for v in mode if key(v) in matched_modes))
    if not ff:
        ok.append('Sem restrição de formato cadastrada')
        notes.append('Formato não documentado na planilha: conferir manual')
    elif not matched_formats:
        bad.append('Nenhum dos formatos solicitados consta nas opções cadastradas')
    else:
        ok.append('Formato(s) compatível(is): '+', '.join(v for v in fmt if key(v) in matched_formats))
    if matched_modes and matched_formats:
        valid=rules[(rules.ID_Teste==tid)&(rules.Modalidade.map(key).isin(matched_modes))&(rules.Formato.map(key).isin(matched_formats))]
        if valid.empty:
            notes.append('Combinação específica modalidade + formato não validada no manual')
        else:
            ok.append('Pelo menos uma combinação modalidade + formato validada no cadastro')
    status=key(t['Situacao_SATEPSI'])
    if status=='desfavoravel':
        bad.append('Situação cadastrada como desfavorável no SATEPSI')
    elif status=='favoravel':
        ok.append('Situação cadastrada como favorável; verificar vigência no SATEPSI')
    else:
        notes.append('Situação favorável atual não confirmada no SATEPSI')
    critical_notes = [n for n in notes if n.startswith(('Sem restrição de idade', 'Escolaridade mínima não interpretada', 'Modalidade não documentada', 'Formato não documentado'))]
    group='Incompatível' if bad else 'Possível — conferir' if critical_notes else 'Compatível com dados cadastrados'
    return {'Teste':clean(t['Nome_Teste']),'Grupo':group,'Atendidos':ok,'Conferir':notes,'Incompatíveis':bad,'Construtos':cc,'Observações':clean(t['Observacoes_Original']),'Fonte':clean(t['URL_Fonte']),'Data':clean(t['Data_Conferencia'])}

if submit:
    if not mode or not fmt:
        st.warning('Selecione pelo menos uma modalidade e um formato de aplicação.')
    elif not wanted:
        st.warning('Selecione pelo menos um construto.')
    else:
        st.subheader('Resultados da busca')
        st.caption('A escolaridade em branco ou indicada como independente não exclui o teste. Alertas sobre validação normativa e condições de aplicação são separados da compatibilidade de escolaridade.')
        st.write(f'Busca para pessoa de **{age} anos**, escolaridade **{school}**, aplicação **{', '.join(mode).lower()}**, formato **{', '.join(fmt).lower()}**, construto(s) **{", ".join(wanted)}**. Aceita outros construtos: **{extra.lower()}**.')
        results=[evaluate(t) for _,t in tests.iterrows()]
        for group in ['Compatível com dados cadastrados','Possível — conferir','Incompatível']:
            subset=[r for r in results if r['Grupo']==group]
            st.markdown(f'### {group} ({len(subset)})')
            if not subset: st.caption('Nenhum instrumento nesta categoria.')
            for r in subset:
                with st.expander(r['Teste']):
                    st.markdown('**Por que apareceu nesta categoria?**')
                    for s in r['Atendidos']: st.write('✓ '+s)
                    for s in r['Conferir']: st.warning('Conferir: '+s)
                    for s in r['Incompatíveis']: st.error('Incompatibilidade: '+s)
                    st.write('**Construtos cadastrados:** '+(', '.join(r['Construtos']) or 'Não informado'))
                    if r['Observações']: st.info('**Informações adicionais / observações da base:** '+r['Observações'])
                    st.caption('Fonte: '+(r['Fonte'] or 'Não informada')+' | Data de conferência: '+(r['Data'] or 'Não informada'))
                    st.link_button('Conferir situação no SATEPSI', SATEPSI)
        output=pd.DataFrame([{'Teste':r['Teste'],'Classificação':r['Grupo'],'Atendidos':'; '.join(r['Atendidos']),'A conferir':'; '.join(r['Conferir']),'Incompatibilidades':'; '.join(r['Incompatíveis']),'Observações':r['Observações']} for r in results])
        st.download_button('Baixar análise (CSV)',output.to_csv(index=False).encode('utf-8-sig'),'analise_testes.csv','text/csv')

with st.expander('📚 Consultar todos os testes (sem filtros)'):
    # Visão integral: colunas originais e opções vinculadas às demais abas.
    overview = tests.copy()
    for sheet, value_col, target in [
        (cons, 'Construto_Padronizado', 'Construtos_cadastrados'),
        (mods, 'Modalidade', 'Modalidades_cadastradas'),
        (forms, 'Formato', 'Formatos_cadastrados')
    ]:
        if value_col in sheet.columns and 'ID_Teste' in sheet.columns:
            grouped = sheet.groupby('ID_Teste')[value_col].apply(
                lambda values: '; '.join(dict.fromkeys(clean(v) for v in values if clean(v)))
            )
            overview[target] = overview['ID_Teste'].map(grouped).fillna('')
        else:
            overview[target] = ''
    st.dataframe(overview, hide_index=True, width='stretch')
    st.caption('Todas as colunas da aba Testes, mais construtos, modalidades e formatos cadastrados nas abas relacionadas.')

st.divider()
st.subheader('Responsabilidade técnica e ética')
st.warning('**Atenção:** esta plataforma é um apoio informativo e não dispensa a análise, a responsabilidade técnica e o julgamento do(a) profissional de Psicologia. Antes de escolher ou aplicar qualquer instrumento, consulte sua situação vigente no SATEPSI, o manual atualizado, as normas de interpretação, o público-alvo e as condições autorizadas de aplicação.')
st.markdown('''**Boas práticas recomendadas:**
- Observar o Código de Ética Profissional do Psicólogo e as resoluções vigentes do Conselho Federal de Psicologia sobre avaliação psicológica e uso de testes.
- Selecionar instrumentos tecnicamente fundamentados, pertinentes à demanda e adequados à pessoa avaliada e ao contexto.
- Respeitar os limites de aplicação, interpretação, padronização e normas do manual; não inferir autorização a partir de campos vazios.
- Garantir sigilo profissional, proteção de dados pessoais e uso responsável dos resultados, em conformidade com as normas aplicáveis.
- Não interpretar uma indicação automática como diagnóstico ou decisão profissional; considerar múltiplas fontes de informação quando cabível.
''')
st.info('**Sobre a atualização:** este protótipo não possui integração direta e automática com o SATEPSI. Como não foi disponibilizado aqui um acesso oficial de consulta automatizada, recomenda-se que o profissional realize a conferência diretamente no portal do CFP. A situação de um teste pode mudar após a elaboração da planilha.')
st.link_button('Abrir SATEPSI para conferência oficial', SATEPSI)
st.caption('Projeto acadêmico. As descrições dos construtos são orientativas e não reproduzem nem substituem os manuais técnicos.')
