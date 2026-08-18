# -*- coding: utf-8 -*-
"""
Contagem Volumetrica - Word -> Excel
=====================================

Le a tabela de contagem digitada no Word (sequencias de digitos, um digito por
veiculo), conta por tipo e preenche as colunas Leve / Onibus / Caminhao / Moto
da planilha, uma aba por movimento.

Codigos:  1 = Moto   2 = Leve   3 = Onibus   4 = Caminhao   5 = Articulado
O articulado (5) e somado dentro de Caminhao.

Nao usa nenhuma biblioteca externa - so a biblioteca padrao do Python.
A planilha original nunca e alterada: o resultado sai em uma copia, com todas
as formulas, imagens e formatacao preservadas.

Uso:
    python main.py              -> abre a interface grafica
    python main.py --cli word.docx plan.xlsx 2,3 saida.xlsx
"""

import os
import re
import shutil
import sys
import zipfile
from collections import Counter, OrderedDict

# ---------------------------------------------------------------------------
# Configuracao
# ---------------------------------------------------------------------------

#: digito digitado no Word -> nome da coluna na planilha
CODIGOS = {
    "1": "Moto",
    "2": "Leve",
    "3": "Onibus",
    "4": "Caminhao",
    "5": "Caminhao",   # articulado somado em caminhao
}

#: nome interno -> como o cabecalho aparece na planilha (com acento)
ROTULOS = {
    "Leve": "Leve",
    "Onibus": "Ônibus",
    "Caminhao": "Caminhão",
    "Moto": "Moto",
}

ORDEM_TIPOS = ["Leve", "Onibus", "Caminhao", "Moto"]

RE_HORA = re.compile(r"^\s*\d{1,2}\s*:\s*\d{2}\s*-\s*\d{1,2}\s*:\s*\d{2}\s*$")


def norm_hora(texto):
    """'13 :45 - 14:00' e '13:45-14:00' viram a mesma chave."""
    return re.sub(r"\s+", "", texto or "")


def norm_rotulo(texto):
    """Compara cabecalhos ignorando acento, caixa e espacos."""
    if not texto:
        return ""
    t = texto.strip().lower()
    for de, para in (
        ("á", "a"), ("à", "a"), ("ã", "a"), ("â", "a"),
        ("é", "e"), ("ê", "e"), ("í", "i"),
        ("ó", "o"), ("ô", "o"), ("õ", "o"),
        ("ú", "u"), ("ç", "c"),
    ):
        t = t.replace(de, para)
    return t


class ErroDeUso(Exception):
    """Erro previsto, com mensagem legivel para o usuario."""


# ---------------------------------------------------------------------------
# Leitura do Word
# ---------------------------------------------------------------------------

W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def _texto_da_celula(tc):
    """Texto de uma celula do Word, com '\\n' entre paragrafos e quebras."""
    import xml.etree.ElementTree as ET

    paragrafos = []
    for p in tc.iter(W_NS + "p"):
        partes = []
        for no in p.iter():
            if no.tag == W_NS + "t":
                partes.append(no.text or "")
            elif no.tag in (W_NS + "br", W_NS + "cr"):
                partes.append("\n")
            elif no.tag == W_NS + "tab":
                partes.append(" ")
        paragrafos.append("".join(partes))
    return "\n".join(paragrafos)


def ler_word(caminho):
    """
    Devolve:
        horas    - lista de horarios na ordem em que aparecem
        linhas   - {hora_normalizada: [texto_col_1, texto_col_2, ...]}
        n_cols   - quantas colunas de dados existem depois da coluna HORA
    """
    import xml.etree.ElementTree as ET

    try:
        with zipfile.ZipFile(caminho) as z:
            xml = z.read("word/document.xml")
    except (zipfile.BadZipFile, KeyError):
        raise ErroDeUso(
            "Arquivo do Word inválido ou corrompido.\n"
            "Ele precisa ser .docx (não .doc)."
        )

    raiz = ET.fromstring(xml)
    tabelas = list(raiz.iter(W_NS + "tbl"))
    if not tabelas:
        raise ErroDeUso("Nenhuma tabela encontrada no Word.")

    # Le todas as tabelas; a de contagem e a que tem mais linhas de horario.
    melhor = None
    for tbl in tabelas:
        grade = []
        for tr in tbl.findall(W_NS + "tr"):
            grade.append([_texto_da_celula(tc) for tc in tr.findall(W_NS + "tc")])
        pontos = sum(
            1 for lin in grade for c in lin if RE_HORA.match(c or "")
        )
        if melhor is None or pontos > melhor[0]:
            melhor = (pontos, grade)

    pontos, grade = melhor
    if pontos == 0:
        raise ErroDeUso(
            "Não encontrei nenhuma linha de horário no Word.\n"
            "Esperado algo como '6:00 - 6:15' na primeira coluna."
        )

    # Qual indice de coluna concentra os horarios?
    votos = Counter()
    for lin in grade:
        for i, c in enumerate(lin):
            if RE_HORA.match(c or ""):
                votos[i] += 1
    col_hora = votos.most_common(1)[0][0]

    horas = []
    linhas = OrderedDict()
    largura = 0
    for lin in grade:
        if len(lin) <= col_hora:
            continue
        bruto = lin[col_hora]
        if not RE_HORA.match(bruto or ""):
            continue
        dados = [c.strip() for c in lin[col_hora + 1:]]
        largura = max(largura, len(dados))
        chave = norm_hora(bruto)
        horas.append(bruto.strip())
        linhas[chave] = dados

    # Descarta colunas totalmente vazias no fim
    n_cols = 0
    for i in range(largura):
        if any(len(d) > i and d[i] for d in linhas.values()):
            n_cols = i + 1
    for chave, d in linhas.items():
        linhas[chave] = (d + [""] * n_cols)[:n_cols]

    if n_cols == 0:
        raise ErroDeUso("O Word não tem nenhuma sequência digitada.")

    return horas, linhas, n_cols, detectar_marcadores(horas, linhas, n_cols)


def detectar_marcadores(horas, linhas, n_cols):
    """
    A primeira celula preenchida de cada coluna e o numero do movimento,
    escrito logo antes do inicio da contagem.

    Devolve [(hora_normalizada, numero) ou None] - um item por coluna.
    Aqui e so um palpite; quem confirma e a lista de movimentos da planilha.
    """
    marcadores = []
    for i in range(n_cols):
        cheias = [(norm_hora(h), linhas[norm_hora(h)][i].strip())
                  for h in horas if linhas[norm_hora(h)][i].strip()]
        candidato = None
        if cheias:
            chave, primeiro = cheias[0]
            if primeiro.isdigit() and len(primeiro) <= 2:
                candidato = (chave, primeiro)
        marcadores.append(candidato)
    return marcadores


def contar(sequencia):
    """
    Conta os digitos de uma celula.
    Devolve (contagens, invalidos) - invalidos e a lista de caracteres estranhos.
    """
    contagens = {t: 0 for t in ORDEM_TIPOS}
    invalidos = []
    for ch in sequencia:
        if ch in " \t\n\r":
            continue
        tipo = CODIGOS.get(ch)
        if tipo is None:
            invalidos.append(ch)
        else:
            contagens[tipo] += 1
    return contagens, invalidos


# ---------------------------------------------------------------------------
# Leitura da planilha
# ---------------------------------------------------------------------------

RE_CELULA = re.compile(
    r'<c r="([A-Z]+)(\d+)"((?:\s[^>/]*)?)(?:/>|>(.*?)</c>)', re.S
)
RE_LINHA = re.compile(r'<row[^>]*\sr="(\d+)"[^>]*(?:/>|>(.*?)</row>)', re.S)


def col_para_num(letras):
    n = 0
    for ch in letras:
        n = n * 26 + (ord(ch) - 64)
    return n


def num_para_col(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _desescapar(txt):
    return (
        txt.replace("&lt;", "<").replace("&gt;", ">")
        .replace("&quot;", '"').replace("&apos;", "'")
        .replace("&amp;", "&")
    )


def _ler_shared_strings(z):
    try:
        xml = z.read("xl/sharedStrings.xml").decode("utf-8")
    except KeyError:
        return []
    itens = []
    for si in re.findall(r"<si>(.*?)</si>", xml, re.S):
        itens.append(_desescapar("".join(re.findall(r"<t[^>]*>(.*?)</t>", si, re.S))))
    return itens


def _valores_da_aba(xml, shared):
    """{ 'C33': texto } - so celulas com conteudo (formula conta como conteudo)."""
    valores = {}
    formulas = set()
    for m in RE_CELULA.finditer(xml):
        ref = m.group(1) + m.group(2)
        attrs = m.group(3) or ""
        inner = m.group(4)
        if not inner:
            continue
        if "<f" in inner:
            formulas.add(ref)
        mv = re.search(r"<v>(.*?)</v>", inner, re.S)
        if mv is None:
            mt = re.search(r"<t[^>]*>(.*?)</t>", inner, re.S)
            if mt:
                valores[ref] = _desescapar(mt.group(1))
            continue
        bruto = mv.group(1)
        if 't="s"' in attrs:
            idx = int(bruto)
            valores[ref] = shared[idx] if idx < len(shared) else ""
        else:
            valores[ref] = _desescapar(bruto)
    return valores, formulas


def ler_xlsx(caminho):
    """
    Devolve a lista de abas de movimento, cada uma como um dicionario:
        nome, arquivo, xml, movimento, linha_cab, colunas {tipo: letra},
        horas {hora_norm: numero_da_linha}, valores, estilos {tipo: (cheio, vazio)}
    """
    try:
        z = zipfile.ZipFile(caminho)
    except zipfile.BadZipFile:
        raise ErroDeUso(
            "Arquivo do Excel inválido ou corrompido.\n"
            "Ele precisa ser .xlsx (não .xls)."
        )

    with z:
        shared = _ler_shared_strings(z)
        try:
            wb = z.read("xl/workbook.xml").decode("utf-8")
            rels = z.read("xl/_rels/workbook.xml.rels").decode("utf-8")
        except KeyError:
            raise ErroDeUso(
                "Este arquivo não é uma planilha do Excel.\n"
                "Confira se você não trocou o Word com a planilha."
            )

        alvo = {}
        for m in re.finditer(r'<Relationship Id="([^"]+)"[^>]*Target="([^"]+)"', rels):
            alvo[m.group(1)] = m.group(2)

        abas = []
        for m in re.finditer(r"<sheet\b[^>]*/>", wb):
            tag = m.group(0)
            nome = re.search(r'name="([^"]*)"', tag)
            rid = re.search(r'r:id="([^"]*)"', tag)
            if not nome or not rid:
                continue
            destino = alvo.get(rid.group(1), "")
            if not destino:
                continue
            arq = "xl/" + destino.lstrip("/")
            try:
                xml = z.read(arq).decode("utf-8")
            except KeyError:
                continue
            abas.append({"nome": _desescapar(nome.group(1)), "arquivo": arq, "xml": xml})

    movimentos = []
    for aba in abas:
        valores, formulas = _valores_da_aba(aba["xml"], shared)
        aba["valores"] = valores

        # linha de cabecalho = onde alguma celula vale "HORA"
        linha_cab = None
        for ref, v in valores.items():
            if norm_rotulo(v) == "hora":
                linha_cab = int(re.match(r"[A-Z]+(\d+)", ref).group(1))
                col_hora = re.match(r"([A-Z]+)", ref).group(1)
                break
        if linha_cab is None:
            continue  # nao e aba de movimento

        # cabecalhos Leve / Onibus / Caminhao / Moto
        colunas = {}
        for tipo in ORDEM_TIPOS:
            alvo_rot = norm_rotulo(ROTULOS[tipo])
            for n in range(col_para_num(col_hora) + 1, col_para_num(col_hora) + 12):
                ref = num_para_col(n) + str(linha_cab)
                if norm_rotulo(valores.get(ref, "")) == alvo_rot:
                    colunas[tipo] = num_para_col(n)
                    break
        if len(colunas) < 4:
            continue

        # numero do movimento: celula a direita do rotulo "Movimento"
        movimento = None
        for ref, v in valores.items():
            if norm_rotulo(v) == "movimento":
                letras = re.match(r"([A-Z]+)", ref).group(1)
                linha = re.match(r"[A-Z]+(\d+)", ref).group(1)
                viz = valores.get(num_para_col(col_para_num(letras) + 1) + linha, "")
                try:
                    movimento = str(int(float(viz)))
                except (TypeError, ValueError):
                    movimento = viz.strip() or None
                break
        if movimento is None:
            movimento = aba["nome"]

        # horarios da coluna HORA
        horas = OrderedDict()
        linha = linha_cab + 1
        vazias = 0
        while vazias < 5:
            v = valores.get(col_hora + str(linha), "")
            if RE_HORA.match(v or ""):
                horas[norm_hora(v)] = linha
                vazias = 0
            else:
                vazias += 1
            linha += 1
        if not horas:
            continue

        # estilo usado em celulas ja preenchidas x vazias, por coluna
        estilos = {}
        primeira, ultima = min(horas.values()), max(horas.values())
        for tipo, letra in colunas.items():
            cheio, vazio = Counter(), Counter()
            for m in RE_CELULA.finditer(aba["xml"]):
                if m.group(1) != letra:
                    continue
                n = int(m.group(2))
                if n < primeira or n > ultima:
                    continue
                est = re.search(r's="(\d+)"', m.group(3) or "")
                est = est.group(1) if est else None
                (cheio if m.group(4) else vazio)[est] += 1
            estilos[tipo] = (
                cheio.most_common(1)[0][0] if cheio else None,
                vazio.most_common(1)[0][0] if vazio else None,
            )

        aba.update(
            movimento=str(movimento),
            linha_cab=linha_cab,
            col_hora=col_hora,
            colunas=colunas,
            horas=horas,
            estilos=estilos,
        )
        movimentos.append(aba)

    if not movimentos:
        raise ErroDeUso(
            "Nenhuma aba de movimento encontrada na planilha.\n"
            "Esperado uma aba com 'HORA' e as colunas Leve / Ônibus / Caminhão / Moto."
        )
    return movimentos


# ---------------------------------------------------------------------------
# Escrita: cirurgia no XML, preservando formulas, imagens e formatacao
# ---------------------------------------------------------------------------

def _escrever_celula(xml_linha, ref, valor, estilo):
    """Substitui (ou cria) a celula `ref` dentro do XML de uma linha."""
    letras = re.match(r"([A-Z]+)", ref).group(1)
    padrao = re.compile(r'<c r="%s"((?:\s[^>/]*)?)(?:/>|>(.*?)</c>)' % ref, re.S)
    m = padrao.search(xml_linha)

    if m:
        attrs = m.group(1) or ""
        attrs = re.sub(r'\st="[^"]*"', "", attrs)          # deixa de ser texto
        if estilo is not None:
            if re.search(r'\ss="\d+"', attrs):
                attrs = re.sub(r'\ss="\d+"', ' s="%s"' % estilo, attrs)
            else:
                attrs += ' s="%s"' % estilo
        nova = '<c r="%s"%s><v>%d</v></c>' % (ref, attrs, valor)
        return xml_linha[: m.start()] + nova + xml_linha[m.end():]

    # celula nao existe: insere na posicao correta
    est = ' s="%s"' % estilo if estilo is not None else ""
    nova = '<c r="%s"%s><v>%d</v></c>' % (ref, est, valor)
    destino = col_para_num(letras)
    for outra in RE_CELULA.finditer(xml_linha):
        if col_para_num(outra.group(1)) > destino:
            return xml_linha[: outra.start()] + nova + xml_linha[outra.start():]
    fim = xml_linha.rfind("</row>")
    if fim == -1:
        return xml_linha
    return xml_linha[:fim] + nova + xml_linha[fim:]


def aplicar_edicoes(xml, edicoes):
    """edicoes: {linha: [(ref, valor, estilo), ...]}"""
    if not edicoes:
        return xml
    saida = []
    fim_anterior = 0
    for m in RE_LINHA.finditer(xml):
        n = int(m.group(1))
        if n not in edicoes:
            continue
        bloco = m.group(0)
        if bloco.endswith("/>"):  # linha vazia auto-fechada
            bloco = bloco[:-2] + "></row>"
        for ref, valor, estilo in edicoes[n]:
            bloco = _escrever_celula(bloco, ref, valor, estilo)
        saida.append(xml[fim_anterior:m.start()])
        saida.append(bloco)
        fim_anterior = m.end()
    saida.append(xml[fim_anterior:])
    return "".join(saida)


def forcar_recalculo(wb_xml):
    """Faz o Excel recalcular tudo ao abrir (totais, fatores, expansao)."""
    if re.search(r"<calcPr\b[^>]*/>", wb_xml):
        return re.sub(r"<calcPr\b[^>]*/>", '<calcPr fullCalcOnLoad="1"/>', wb_xml, count=1)
    if "<calcPr" in wb_xml:
        return re.sub(r"<calcPr\b", '<calcPr fullCalcOnLoad="1" ', wb_xml, count=1)
    for ancora in ("</definedNames>", "</sheets>"):
        if ancora in wb_xml:
            return wb_xml.replace(ancora, ancora + '<calcPr fullCalcOnLoad="1"/>', 1)
    return wb_xml


def salvar_copia(origem, destino, novos_xml):
    """Reescreve o .xlsx trocando so os arquivos internos alterados."""
    if os.path.abspath(origem) == os.path.abspath(destino):
        raise ErroDeUso("O arquivo de saída precisa ser diferente do original.")
    temporario = destino + ".tmp"
    with zipfile.ZipFile(origem) as zin:
        infos = zin.infolist()
        with zipfile.ZipFile(temporario, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in infos:
                if info.filename in novos_xml:
                    dados = novos_xml[info.filename].encode("utf-8")
                else:
                    dados = zin.read(info.filename)
                novo = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                novo.compress_type = info.compress_type
                novo.external_attr = info.external_attr
                zout.writestr(novo, dados)
    shutil.move(temporario, destino)


# ---------------------------------------------------------------------------
# Processamento
# ---------------------------------------------------------------------------

def processar(caminho_word, caminho_xlsx, destino, mapeamento, sobrescrever=False):
    """
    mapeamento: {indice_da_coluna_do_word: numero_do_movimento}
    Devolve o texto do relatorio.
    """
    horas_word, linhas_word, n_cols, marcadores = ler_word(caminho_word)
    movimentos = ler_xlsx(caminho_xlsx)
    por_numero = {m["movimento"]: m for m in movimentos}

    rel = Relatorio()
    rel.add("RELATÓRIO DA CONTAGEM")
    rel.add("=" * 66)
    rel.add("Word .....: %s" % os.path.basename(caminho_word))
    rel.add("Planilha .: %s" % os.path.basename(caminho_xlsx))
    rel.add("Saída ....: %s" % os.path.basename(destino))
    rel.add("")

    novos_xml = {}
    total_escrito = 0

    for indice in sorted(mapeamento):
        numero = str(mapeamento[indice]).strip()
        aba = por_numero.get(numero)
        if aba is None:
            raise ErroDeUso(
                "Não existe aba para o movimento '%s'.\n"
                "Movimentos disponíveis na planilha: %s"
                % (numero, ", ".join(sorted(por_numero)))
            )

        rel.add("-" * 66)
        rel.add("COLUNA %d DO WORD  ->  MOVIMENTO %s  (aba \"%s\")"
                % (indice + 1, numero, aba["nome"]))
        rel.add("-" * 66)
        rel.contexto = "mov %s" % numero

        # A celula curta antes do inicio da contagem e o rotulo do movimento,
        # nao veiculo. So confirmo como rotulo se o numero existir na planilha.
        marcador = marcadores[indice] if indice < len(marcadores) else None
        rotulo_em = None
        if marcador and marcador[1] in por_numero:
            rotulo_em = marcador[0]
            rel.add("Movimento marcado no Word .: %s" % marcador[1])
            if marcador[1] != numero:
                rel.aviso("o Word marca esta coluna como movimento %s, "
                          "mas voce escolheu o movimento %s" % (marcador[1], numero))

        edicoes = {}
        escritos = ignorados = 0
        soma = {t: 0 for t in ORDEM_TIPOS}

        for hora_txt in horas_word:
            chave = norm_hora(hora_txt)
            if chave == rotulo_em:
                continue
            sequencia = linhas_word[chave][indice]
            if not sequencia.strip():
                continue

            contagens, invalidos = contar(sequencia)
            if invalidos:
                rel.aviso("%s: caractere inválido %s (ignorado)"
                          % (hora_txt, " ".join(sorted(set(invalidos)))))
            if "\n" in sequencia:
                pedacos = [p for p in sequencia.split("\n") if p.strip()]
                rel.aviso("%s: célula com %d blocos separados por quebra de linha "
                          "(%s dígitos) — confira se é tudo do mesmo intervalo"
                          % (hora_txt, len(pedacos),
                             "+".join(str(len(p.strip())) for p in pedacos)))
            total_veiculos = sum(contagens.values())
            if 0 < total_veiculos < 5:
                rel.aviso("%s: só %d veículo(s) no intervalo inteiro — "
                          "parece digitação de teste, confira"
                          % (hora_txt, total_veiculos))

            linha = aba["horas"].get(chave)
            if linha is None:
                rel.erro("%s: preenchido no Word mas não existe na planilha" % hora_txt)
                continue

            # ja preenchido na planilha?
            atual = {}
            ocupada = False
            for tipo in ORDEM_TIPOS:
                ref = aba["colunas"][tipo] + str(linha)
                bruto = aba["valores"].get(ref, "")
                if bruto not in ("", None):
                    ocupada = True
                    try:
                        atual[tipo] = int(float(bruto))
                    except ValueError:
                        atual[tipo] = bruto

            if ocupada:
                difs = [
                    "%s %s->%s" % (ROTULOS[t], atual.get(t, 0), contagens[t])
                    for t in ORDEM_TIPOS if atual.get(t, 0) != contagens[t]
                ]
                if difs:
                    rel.divergencia("%s: %s" % (hora_txt, "  ".join(difs)))
                if not sobrescrever:
                    ignorados += 1
                    continue

            for tipo in ORDEM_TIPOS:
                ref = aba["colunas"][tipo] + str(linha)
                cheio, vazio = aba["estilos"].get(tipo, (None, None))
                edicoes.setdefault(linha, []).append(
                    (ref, contagens[tipo], cheio if cheio is not None else vazio)
                )
                soma[tipo] += contagens[tipo]
            escritos += 1

        if edicoes:
            aba["xml"] = aplicar_edicoes(aba["xml"], edicoes)
            novos_xml[aba["arquivo"]] = aba["xml"]

        # intervalos que existem na planilha mas nao no Word
        faltando = [h for h in aba["horas"] if h not in linhas_word]
        if faltando:
            rel.aviso("%d intervalo(s) da planilha não existem no Word "
                      "(não preenchidos)" % len(faltando))

        rel.add("Intervalos preenchidos ...: %d" % escritos)
        if ignorados:
            rel.add("Já preenchidos, mantidos .: %d" % ignorados)
        rel.add("Total contado ............: %s   (%d veículos)"
                % ("  ".join("%s %d" % (ROTULOS[t], soma[t]) for t in ORDEM_TIPOS),
                   sum(soma.values())))
        rel.add("")
        total_escrito += escritos

    if not novos_xml:
        rel.add("Nada foi escrito: todos os intervalos já estavam preenchidos.")
    else:
        with zipfile.ZipFile(caminho_xlsx) as z:
            wb = z.read("xl/workbook.xml").decode("utf-8")
        novos_xml["xl/workbook.xml"] = forcar_recalculo(wb)
        salvar_copia(caminho_xlsx, destino, novos_xml)

    rel.rodape(total_escrito)
    return rel.texto()


class Relatorio(object):
    def __init__(self):
        self.linhas = []
        self.avisos = []
        self.erros = []
        self.divergencias = []
        self.contexto = ""

    def _marcar(self, txt):
        return ("%s | %s" % (self.contexto, txt)) if self.contexto else txt

    def add(self, txt):
        self.linhas.append(txt)

    def aviso(self, txt):
        self.avisos.append(self._marcar(txt))

    def erro(self, txt):
        self.erros.append(self._marcar(txt))

    def divergencia(self, txt):
        self.divergencias.append(self._marcar(txt))

    def rodape(self, total):
        L = self.linhas
        L.append("=" * 66)
        if self.divergencias:
            L.append("")
            L.append("DIVERGÊNCIAS  (já estava na planilha -> contagem real)")
            L.append("Estes intervalos NÃO foram alterados. Confira um a um:")
            L.append("")
            L.extend("  " + d for d in self.divergencias)
        if self.avisos:
            L.append("")
            L.append("AVISOS")
            L.extend("  " + a for a in self.avisos)
        if self.erros:
            L.append("")
            L.append("ERROS")
            L.extend("  " + e for e in self.erros)
        L.append("")
        L.append("=" * 66)
        L.append("%d intervalos preenchidos | %d divergências | %d avisos | %d erros"
                 % (total, len(self.divergencias), len(self.avisos), len(self.erros)))

    def texto(self):
        return "\n".join(self.linhas)


def movimentos_do_word(caminho_word, caminho_xlsx):
    """Mapeamento coluna -> movimento lido dos rotulos do proprio Word."""
    _, _, n_cols, marcadores = ler_word(caminho_word)
    numeros = {m["movimento"] for m in ler_xlsx(caminho_xlsx)}
    mapeamento = {}
    for i in range(n_cols):
        if marcadores[i] and marcadores[i][1] in numeros:
            mapeamento[i] = marcadores[i][1]
    if len(mapeamento) < n_cols:
        faltando = [str(i + 1) for i in range(n_cols) if i not in mapeamento]
        raise ErroDeUso(
            "Não consegui ler o movimento da(s) coluna(s) %s no Word.\n"
            "Informe os movimentos na ordem das colunas, ex: 2,3"
            % ", ".join(faltando)
        )
    return mapeamento


def nome_de_saida(caminho_xlsx):
    base, ext = os.path.splitext(caminho_xlsx)
    return base + " - preenchido" + ext


# ---------------------------------------------------------------------------
# Interface grafica
# ---------------------------------------------------------------------------

def abrir_interface():
    try:
        import tkinter as tk
        from tkinter import filedialog, messagebox, ttk, scrolledtext
    except ImportError:
        print(
            "A interface gráfica precisa do Tkinter, que não está instalado.\n"
            "No Windows e no Mac ele já vem com o Python.\n"
            "No Linux:  sudo apt install python3-tk\n\n"
            "Enquanto isso dá para usar pelo terminal:\n"
            "  python main.py --cli arquivo.docx planilha.xlsx 2,3"
        )
        return

    estado = {"word": None, "xlsx": None, "dados": None, "movs": None,
              "campos": [], "mapa": None}

    janela = tk.Tk()
    janela.title("Contagem Volumétrica")
    janela.geometry("760x620")
    janela.minsize(680, 520)

    corpo = ttk.Frame(janela, padding=14)
    corpo.pack(fill="both", expand=True)

    # --- passo 1: arquivos -------------------------------------------------
    caixa = ttk.LabelFrame(corpo, text=" Arquivos ", padding=10)
    caixa.pack(fill="x")
    caixa.columnconfigure(1, weight=1)

    lbl_word = ttk.Label(caixa, text="nenhum arquivo escolhido", foreground="#888")
    lbl_xlsx = ttk.Label(caixa, text="nenhum arquivo escolhido", foreground="#888")

    def escolher_word():
        c = filedialog.askopenfilename(
            title="Escolha o Word da contagem",
            filetypes=[("Documento do Word", "*.docx"), ("Todos", "*.*")],
        )
        if not c:
            return
        try:
            estado["dados"] = ler_word(c)
        except ErroDeUso as e:
            messagebox.showerror("Word", str(e))
            return
        except Exception as e:
            messagebox.showerror("Word", "Não consegui ler o Word:\n\n%s" % e)
            return
        estado["word"] = c
        lbl_word.config(text=os.path.basename(c), foreground="#000")
        montar_mapeamento()

    def escolher_xlsx():
        c = filedialog.askopenfilename(
            title="Escolha a planilha",
            filetypes=[("Planilha do Excel", "*.xlsx"), ("Todos", "*.*")],
        )
        if not c:
            return
        try:
            estado["movs"] = ler_xlsx(c)
        except ErroDeUso as e:
            messagebox.showerror("Planilha", str(e))
            return
        except Exception as e:
            messagebox.showerror("Planilha", "Não consegui ler a planilha:\n\n%s" % e)
            return
        estado["xlsx"] = c
        lbl_xlsx.config(text=os.path.basename(c), foreground="#000")
        montar_mapeamento()

    ttk.Button(caixa, text="Escolher Word...", width=20,
               command=escolher_word).grid(row=0, column=0, sticky="w", pady=3)
    lbl_word.grid(row=0, column=1, sticky="w", padx=10)
    ttk.Button(caixa, text="Escolher planilha...", width=20,
               command=escolher_xlsx).grid(row=1, column=0, sticky="w", pady=3)
    lbl_xlsx.grid(row=1, column=1, sticky="w", padx=10)

    # O movimento vem do proprio Word. Esta caixa so aparece se faltar algum.
    resumo = ttk.Label(caixa, text="", foreground="#555")
    resumo.grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))

    caixa2 = ttk.LabelFrame(corpo, text=" Movimento de cada coluna ", padding=10)
    dica = ttk.Label(caixa2, text="", foreground="#000")
    dica.pack(anchor="w")
    grade = ttk.Frame(caixa2)
    grade.pack(fill="x")

    def montar_mapeamento():
        for w in grade.winfo_children():
            w.destroy()
        estado["campos"] = []
        estado["mapa"] = None
        caixa2.pack_forget()
        resumo.config(text="")
        if not (estado["dados"] and estado["movs"]):
            return
        horas, linhas, n_cols, marcadores = estado["dados"]
        numeros = [m["movimento"] for m in estado["movs"]]

        # caminho normal: o numero esta no Word e existe como aba na planilha
        automatico = {}
        for i in range(n_cols):
            if marcadores[i] and marcadores[i][1] in numeros:
                automatico[i] = marcadores[i][1]
        if len(automatico) == n_cols:
            estado["mapa"] = automatico
            resumo.config(text="Movimentos: " + "   ".join(
                "coluna %d → %s" % (i + 1, automatico[i]) for i in sorted(automatico)))
            return

        # caminho de excecao: faltou algum, pergunta
        caixa2.pack(fill="x", pady=(12, 0), after=caixa)
        faltando = [str(i + 1) for i in range(n_cols) if i not in automatico]
        dica.config(text="Não achei o número do movimento na coluna %s do Word. "
                         "Confira abaixo." % ", ".join(faltando))

        for i in range(n_cols):
            marcado = marcadores[i][1] if (marcadores[i]
                                           and marcadores[i][1] in numeros) else None
            # so desconta a celula do rotulo se ela foi confirmada como rotulo
            cheias = [h for h in horas
                      if linhas[norm_hora(h)][i].strip()
                      and (marcado is None or norm_hora(h) != marcadores[i][0])]
            amostra = ""
            if cheias:
                # mostra o intervalo mais movimentado: e o mais representativo
                hora = max(cheias, key=lambda h: len(linhas[norm_hora(h)][i]))
                seq = linhas[norm_hora(hora)][i].replace("\n", "")
                amostra = "   pico %s: %d veículos  %s%s" % (
                    hora, len(seq), seq[:14], "…" if len(seq) > 14 else "")
            ttk.Label(grade, text="Coluna %d" % (i + 1),
                      font=("TkDefaultFont", 9, "bold")).grid(
                row=i, column=0, sticky="w", pady=4)
            ttk.Label(grade, text="%d intervalos   %s" % (len(cheias), amostra),
                      foreground="#555").grid(row=i, column=1, sticky="w", padx=10)
            ttk.Label(grade,
                      text="movimento" if marcado is None else "no Word:",
                      foreground="#888").grid(row=i, column=2, sticky="e")
            padrao = marcado or (numeros[i] if i < len(numeros) else "")
            var = tk.StringVar(value=padrao)
            combo = ttk.Combobox(grade, textvariable=var, values=numeros,
                                 width=6, state="normal")
            combo.grid(row=i, column=3, sticky="e", padx=(6, 2))
            estado["campos"].append(var)
        grade.columnconfigure(1, weight=1)

    # --- passo 3: rodar ----------------------------------------------------
    barra = ttk.Frame(corpo)
    barra.pack(fill="x", pady=(12, 6))

    sobrescrever = tk.BooleanVar(value=False)
    ttk.Checkbutton(barra, text="Também corrigir intervalos já preenchidos",
                    variable=sobrescrever).pack(side="left")

    saida = scrolledtext.ScrolledText(corpo, height=16, wrap="none",
                                      font=("Courier New", 9))
    saida.pack(fill="both", expand=True)

    def mostrar(txt):
        saida.delete("1.0", "end")
        saida.insert("1.0", txt)

    def rodar():
        if not estado["word"] or not estado["xlsx"]:
            messagebox.showwarning("Falta arquivo",
                                   "Escolha o Word e a planilha primeiro.")
            return
        mapeamento = estado.get("mapa")
        if not mapeamento:
            mapeamento = {}
            vistos = {}
            for i, var in enumerate(estado["campos"]):
                v = var.get().strip()
                if not v:
                    messagebox.showwarning(
                        "Falta o movimento",
                        "Diga o número do movimento da coluna %d." % (i + 1))
                    return
                if v in vistos:
                    messagebox.showwarning(
                        "Movimento repetido",
                        "As colunas %d e %d apontam para o movimento %s."
                        % (vistos[v] + 1, i + 1, v))
                    return
                vistos[v] = i
                mapeamento[i] = v

        destino = filedialog.asksaveasfilename(
            title="Salvar planilha preenchida como",
            defaultextension=".xlsx",
            initialfile=os.path.basename(nome_de_saida(estado["xlsx"])),
            initialdir=os.path.dirname(estado["xlsx"]),
            filetypes=[("Planilha do Excel", "*.xlsx")],
        )
        if not destino:
            return
        try:
            texto = processar(estado["word"], estado["xlsx"], destino,
                              mapeamento, sobrescrever.get())
        except ErroDeUso as e:
            messagebox.showerror("Nao deu para continuar", str(e))
            return
        except Exception as e:
            import traceback
            mostrar(traceback.format_exc())
            messagebox.showerror("Erro inesperado", str(e))
            return
        mostrar(texto)
        estado["relatorio"] = texto
        messagebox.showinfo("Pronto", "Planilha salva em:\n\n%s" % destino)

    def salvar_relatorio():
        if not estado.get("relatorio"):
            messagebox.showwarning("Sem relatório", "Rode a contagem primeiro.")
            return
        c = filedialog.asksaveasfilename(
            title="Salvar relatório", defaultextension=".txt",
            initialfile="relatório da contagem.txt",
            filetypes=[("Texto", "*.txt")])
        if not c:
            return
        with open(c, "w", encoding="utf-8") as f:
            f.write(estado["relatorio"])
        messagebox.showinfo("Pronto", "Relatório salvo.")

    ttk.Button(barra, text="Salvar relatório",
               command=salvar_relatorio).pack(side="right", padx=(8, 0))
    ttk.Button(barra, text="Contar e preencher",
               command=rodar).pack(side="right")

    mostrar("1. Escolha o Word e a planilha.\n"
            "2. Clique em 'Contar e preencher'.\n\n"
            "O movimento de cada coluna vem do número que você escreve no Word,\n"
            "na linha antes do início da contagem.\n\n"
            "A planilha original não é alterada — o resultado sai em uma cópia.\n"
            "Intervalos já preenchidos são mantidos e listados como divergência.")
    janela.mainloop()


# ---------------------------------------------------------------------------

def main():
    if "--cli" in sys.argv:
        args = [a for a in sys.argv[1:] if a != "--cli"]
        if len(args) < 2:
            print("uso: --cli <word.docx> <planilha.xlsx> [movimentos] "
                  "[saida.xlsx] [--sobrescrever]")
            print("     sem <movimentos>, usa o numero marcado no proprio Word")
            return 2
        sobrescrever = "--sobrescrever" in args
        args = [a for a in args if a != "--sobrescrever"]
        word, xlsx = args[0], args[1]
        extras = args[2:]
        movs = next((a for a in extras if not a.lower().endswith(".xlsx")), None)
        destino = next((a for a in extras if a.lower().endswith(".xlsx")),
                       nome_de_saida(xlsx))
        try:
            if movs:
                mapeamento = {i: m.strip() for i, m in enumerate(movs.split(","))}
            else:
                mapeamento = movimentos_do_word(word, xlsx)
        except ErroDeUso as e:
            print("ERRO: %s" % e)
            return 1
        try:
            print(processar(word, xlsx, destino, mapeamento, sobrescrever))
        except ErroDeUso as e:
            print("ERRO: %s" % e)
            return 1
        return 0
    abrir_interface()
    return 0


if __name__ == "__main__":
    sys.exit(main())
