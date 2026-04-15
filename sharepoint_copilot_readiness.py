#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lê a planilha Excel de permissões SharePoint e gera um relatório DOCX
de PREPARAÇÃO PARA O MICROSOFT 365 COPILOT.

Foco: prevenir vazamento de informação via Copilot.

Analisa 15 riscos específicos do Copilot:
  CR-001  Grupos amplos que o Copilot pode explorar (Everyone, Todos, etc.)
  CR-002  Utilizadores com acesso excessivo (Full Control onde não deviam)
  CR-003  Bibliotecas sem sensitivity label (Copilot acede tudo)
  CR-004  Pastas com herança quebrada expõem dados inesperados
  CR-005  Permissões diretas a utilizadores (difícil rastrear)
  CR-006  Conteúdo acessível a demasiadas pessoas (oversharing)
  CR-007  Sites candidatos a Restricted Content Discovery (RCD)
  CR-008  Bibliotecas candidatas a divisão (isolamento de dados)
  CR-009  Permissões de edição no site raiz (propaga para tudo)
  CR-010  Níveis de permissão que permitem partilha
  CR-011  Contas órfãs com acesso ativo (risco silencioso)
  CR-012  Inconsistência de acesso (utilizador vê dados imprevistos)
  CR-013  Profundidade de permissões (Copilot expõe pastas escondidas)
  CR-014  Cenários de prompt injection (dados misturados)
  CR-015  Checklist de Governance Copilot-Ready

Uso:
    pip install openpyxl python-docx matplotlib
    python sharepoint_copilot_readiness.py "PermissoesSharePoint.xlsx"
    python sharepoint_copilot_readiness.py "PermissoesSharePoint.xlsx" -o copilot_readiness.docx --abrir
"""

import argparse
import io
import os
import re
import sys
from collections import defaultdict, Counter
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("ERRO: pip install openpyxl"); sys.exit(1)
try:
    from docx import Document
    from docx.shared import Inches, Pt, Cm, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.oxml.ns import nsdecls
    from docx.oxml import parse_xml
except ImportError:
    print("ERRO: pip install python-docx"); sys.exit(1)

try:
    import matplotlib; matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    HAS_MPL = True
except ImportError:
    HAS_MPL = False


# ============================================================
# Cores e constantes
# ============================================================
C_BLUE = RGBColor(0x1E, 0x3A, 0x5F)
C_DARK = RGBColor(0x0F, 0x17, 0x2A)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_TEXT = RGBColor(0x33, 0x33, 0x33)
C_TEXT2 = RGBColor(0x64, 0x74, 0x8B)
C_RED = RGBColor(0xEF, 0x44, 0x44)
C_ORANGE = RGBColor(0xF9, 0x7E, 0x16)
C_YELLOW = RGBColor(0xF5, 0x9E, 0x0B)
C_GREEN = RGBColor(0x10, 0xB9, 0x81)
C_BLUE2 = RGBColor(0x3B, 0x82, 0xF6)
C_PURPLE = RGBColor(0x8B, 0x5C, 0xF6)

SEV_C = {"critico": C_RED, "alto": C_ORANGE, "medio": C_YELLOW, "baixo": C_BLUE2, "ok": C_GREEN}
SEV_L = {"critico": "CRÍTICO", "alto": "ALTO", "medio": "MÉDIO", "baixo": "BAIXO", "ok": "OK"}
SEV_I = {"critico": "🔴", "alto": "🟠", "medio": "🟡", "baixo": "🔵", "ok": "✅"}
SEV_ORDER = {"critico": 0, "alto": 1, "medio": 2, "baixo": 3, "ok": 4}


# ============================================================
# Leitura Excel + normalização
# ============================================================
def read_excel(fp):
    if not os.path.isfile(fp): print(f"ERRO: {fp} não encontrado"); sys.exit(1)
    wb = openpyxl.load_workbook(fp, read_only=True, data_only=True)
    data = {"resumo": {}, "site": [], "bibliotecas": [], "subpastas": [], "consolidado": []}
    sm = {}
    for n in wb.sheetnames:
        lo = n.lower().strip()
        if lo == "resumo": sm["resumo"] = n
        elif lo == "site": sm["site"] = n
        elif lo in ("bibliotecas", "libraries"): sm["bibliotecas"] = n
        elif lo in ("subpastas", "folders", "subfolders"): sm["subpastas"] = n
        elif lo in ("consolidado", "consolidated", "all"): sm["consolidado"] = n
    if "resumo" in sm:
        for row in wb[sm["resumo"]].iter_rows(min_row=2, values_only=True):
            if row and row[0]: data["resumo"][str(row[0]).strip()] = row[1] if len(row) > 1 else ""
    def rs(k):
        if k not in sm: return []
        rows = list(wb[sm[k]].iter_rows(values_only=True))
        if not rows: return []
        hdr = [str(h).strip().lower() if h else f"c{i}" for i, h in enumerate(rows[0])]
        recs = []
        for r in rows[1:]:
            if not r or all(c is None for c in r): continue
            recs.append({hdr[i]: str(v).strip() if v is not None else "" for i, v in enumerate(r) if i < len(hdr)})
        return recs
    for k in ["site", "bibliotecas", "subpastas", "consolidado"]: data[k] = rs(k)
    if not data["site"] and not data["bibliotecas"] and data["consolidado"]:
        for rec in data["consolidado"]:
            n = norm(rec); e = n.get("escopo", "").lower()
            if e == "site": data["site"].append(rec)
            elif e in ("biblioteca", "library"): data["bibliotecas"].append(rec)
            elif e in ("pasta", "folder"): data["subpastas"].append(rec)
    wb.close(); return data

FM = {
    "escopo": "escopo", "scope": "escopo", "recurso": "recurso", "resource": "recurso",
    "caminhocompleto": "caminho", "caminho completo": "caminho", "caminho": "caminho",
    "path": "caminho", "fullpath": "caminho", "full path": "caminho",
    "permissoesunicas": "unicas", "permissões únicas": "unicas", "permissoesúnicas": "unicas",
    "unique": "unicas", "hasuniqueroleassignments": "unicas", "uniquepermissions": "unicas",
    "permissões unicas": "unicas",
    "tipomembro": "tipo_membro", "tipo membro": "tipo_membro", "membertype": "tipo_membro",
    "member type": "tipo_membro", "tipo": "tipo_membro",
    "membro": "membro", "member": "membro", "title": "membro",
    "loginmembro": "login", "login membro": "login", "login": "login",
    "loginname": "login", "memberlogin": "login",
    "niveispermissao": "niveis", "níveis permissão": "niveis", "niveispermissão": "niveis",
    "niveis permissao": "niveis", "níveispermissão": "niveis", "permissionlevels": "niveis",
    "permission levels": "niveis", "roles": "niveis", "role": "niveis",
    "niveis de permissao": "niveis", "níveis de permissão": "niveis",
}

def norm(rec):
    out = {}
    for k, v in rec.items():
        c = k.lower().strip().replace("_", "").replace("-", "")
        m = FM.get(c) or FM.get(k.lower().strip())
        out[m if m else c] = v
    return out

def is_uniq(v): return str(v).lower().strip() in ("true", "sim", "yes", "1", "verdadeiro")
def is_user(t): return "user" in t.lower() or t.lower() == "1"
def is_fc(n): nl = n.lower(); return "full control" in nl or "controlo total" in nl or "controle total" in nl
def is_edit(n): nl = n.lower(); return any(x in nl for x in ["edit", "contribute", "edição", "contribui"])
def is_read(n):
    nl = n.lower(); ls = [l.strip() for l in nl.split(";") if l.strip()]
    return bool(ls) and all(any(r in l for r in ["read", "leitura", "view", "apenas"]) for l in ls)


# ============================================================
# Motor de Auditoria Copilot
# ============================================================
class CopilotAuditEngine:
    def __init__(self, data):
        self.data = data; self.recs = []; self.findings = []; self.stats = {}
        for sk in ["site", "bibliotecas", "subpastas"]:
            for rec in data[sk]:
                n = norm(rec); n["_sk"] = sk; self.recs.append(n)

    def run(self):
        self._stats()
        self._cr001(); self._cr002(); self._cr003(); self._cr004(); self._cr005()
        self._cr006(); self._cr007(); self._cr008(); self._cr009(); self._cr010()
        self._cr011(); self._cr012(); self._cr013(); self._cr014(); self._cr015()
        self.findings.sort(key=lambda f: SEV_ORDER.get(f["sev"], 5))
        return self.findings, self.stats

    def _add(self, **kw): self.findings.append(kw)

    def _stats(self):
        t = len(self.recs)
        uc = sum(1 for r in self.recs if is_uniq(r.get("unicas", "")))
        mems = set(r.get("membro", "") for r in self.recs if r.get("membro"))
        du = set(r.get("membro", "") for r in self.recs if is_user(r.get("tipo_membro", "")))
        grps = mems - du
        fc = sum(1 for r in self.recs if is_fc(r.get("niveis", "")))
        # Membros com acesso amplo (aparecem em 5+ locais)
        mem_count = Counter(r.get("membro", "") for r in self.recs if r.get("membro"))
        wide_access = {m for m, c in mem_count.items() if c >= 5}
        # Locais acessíveis por 10+ membros
        path_mems = defaultdict(set)
        for r in self.recs:
            p = r.get("caminho", r.get("recurso", ""))
            if p and r.get("membro"): path_mems[p].add(r.get("membro", ""))
        overshared = {p for p, ms in path_mems.items() if len(ms) >= 10}
        # Libs
        libs = set(r.get("recurso", "") for r in self.recs if r.get("_sk") == "bibliotecas")
        # Libs que precisam de split
        lfs = self._lib_folder_data()
        ls = sum(1 for lib, recs in lfs.items()
                 if len(set(r.get("caminho", "").split("/")[1] for r in recs
                           if len(r.get("caminho", "").split("/")) >= 2)) >= 3)

        score = 100
        ur = uc / max(t, 1)
        if ur > 0.5: score -= 20
        elif ur > 0.3: score -= 12
        elif ur > 0.1: score -= 5
        dr = len(du) / max(len(mems), 1)
        if dr > 0.5: score -= 18
        elif dr > 0.3: score -= 10
        fr = fc / max(t, 1)
        if fr > 0.4: score -= 15
        elif fr > 0.2: score -= 8
        if len(wide_access) > 10: score -= 10
        elif len(wide_access) > 5: score -= 5
        if len(overshared) > 10: score -= 10
        elif len(overshared) > 3: score -= 5
        if ls > 2: score -= 10
        elif ls > 0: score -= 5
        score = max(0, min(100, score))

        self.stats = {
            "total": t, "unique_count": uc, "unique_ratio": round(uc / max(t, 1) * 100, 1),
            "total_members": len(mems), "direct_users": len(du), "groups": len(grps),
            "direct_user_ratio": round(len(du) / max(len(mems), 1) * 100, 1),
            "fc_count": fc, "fc_ratio": round(fc / max(t, 1) * 100, 1),
            "wide_access_members": len(wide_access), "overshared_paths": len(overshared),
            "lib_count": len(libs), "libs_split": ls, "copilot_score": score,
            "du_names": sorted(du), "grp_names": sorted(grps),
            "wide_access_names": sorted(wide_access), "overshared_list": sorted(overshared),
        }

    def _lib_folder_data(self):
        lf = defaultdict(list)
        for r in self.recs:
            if r.get("_sk") == "subpastas" and is_uniq(r.get("unicas", "")):
                p = r.get("caminho", ""); parts = p.split("/")
                lf[parts[0] if parts else "?"].append(r)
        return lf

    # ---- CR-001: Grupos amplos ----
    def _cr001(self):
        pats = ["everyone", "todos", "all users", "todos os utilizadores",
                "everyone except external", "nt authority", "all authenticated"]
        found = []
        for r in self.recs:
            cb = (r.get("membro", "") + " " + r.get("login", "")).lower()
            for p in pats:
                if p in cb:
                    path = r.get("caminho", r.get("recurso", ""))
                    niv = r.get("niveis", "")
                    risk = "CRÍTICO" if is_fc(niv) or is_edit(niv) else "Alto"
                    found.append(f"[{risk}] {r.get('membro', '')} → {niv} em {path}")
                    break
        has_fc = any("CRÍTICO" in f for f in found)
        sev = "critico" if has_fc else "alto" if len(found) > 3 else "medio" if found else "ok"
        self._add(id="CR-001", title="Grupos Amplos Acessíveis pelo Copilot", cat="Vazamento Direto", sev=sev,
            copilot_risk="O Copilot respeita permissões do utilizador. Se um utilizador pertence a 'Everyone', o Copilot pode surfacear QUALQUER documento do site nas respostas.",
            scenario="Utilizador pergunta: 'Quais são os salários da equipa?' → Copilot encontra planilha de RH partilhada com Everyone e apresenta os dados.",
            desc=f"{len(found)} atribuições a grupos amplos encontradas." + (" INCLUI FULL CONTROL/EDIT — risco máximo." if has_fc else ""),
            rec="AÇÃO IMEDIATA (antes de ativar Copilot):\n• Remover TODOS os grupos 'Everyone'/'Todos' do SharePoint.\n• Substituir por grupos Azure AD específicos.\n• Se impossível remover, rebaixar para Read-only no mínimo.\n• Aplicar Restricted Content Discovery (RCD) nos sites afetados:\n  Set-SPOSite -Identity <url> -RestrictContentOrgWideSearch $true",
            affected=found[:20], metric=str(len(found)), mlabel="Atribuições amplas", thr="0 para Copilot")

    # ---- CR-002: Full Control excessivo ----
    def _cr002(self):
        fm = defaultdict(list)
        for r in self.recs:
            if is_fc(r.get("niveis", "")):
                fm[r.get("membro", "")].append(r.get("caminho", r.get("recurso", "")))
        fc = self.stats["fc_count"]; fr = self.stats["fc_ratio"]
        sev = "critico" if fr > 40 else "alto" if fr > 20 else "medio" if fr > 10 else "baixo" if fc > 0 else "ok"
        aff = [f"{m} → {len(p)} locais com Full Control" for m, p in sorted(fm.items(), key=lambda x: -len(x[1]))][:15]
        self._add(id="CR-002", title="Full Control Expõe Capacidade de Partilha", cat="Escalação de Privilégio", sev=sev,
            copilot_risk="Utilizadores com Full Control podem alterar permissões. Se o Copilot sugere 'partilhar este documento', o utilizador com FC pode acidentalmente expor dados a toda a organização.",
            scenario="Utilizador com FC pede: 'Ajuda-me a partilhar este relatório financeiro com a equipa' → Copilot facilita a partilha com permissões excessivas.",
            desc=f"{fc} atribuições Full Control ({fr}%), {len(fm)} membros.",
            rec="• Rebaixar Full Control para 'Edit' ou 'Contribute' em bibliotecas/pastas.\n• Reservar Full Control APENAS para Site Collection Admins.\n• Antes do Copilot: auditar quem tem FC e porquê.\n• Configurar DLP no Purview para bloquear partilha de conteúdo sensível.",
            affected=aff, metric=f"{fr}%", mlabel="Full Control", thr="< 5% para Copilot")

    # ---- CR-003: Sem sensitivity labels ----
    def _cr003(self):
        libs = set(r.get("recurso", "") for r in self.recs if r.get("_sk") == "bibliotecas")
        sev = "alto" if len(libs) > 3 else "medio" if libs else "ok"
        self._add(id="CR-003", title="Bibliotecas Sem Sensitivity Labels (Purview)", cat="Classificação", sev=sev,
            copilot_risk="Sem sensitivity labels, o Copilot trata TODO o conteúdo como 'General'. Documentos confidenciais ficam ao mesmo nível que documentos públicos — o Copilot pode misturá-los nas respostas.",
            scenario="Utilizador pergunta: 'Resume os documentos recentes' → Copilot inclui contratos confidenciais, dados pessoais e documentos internos sem distinção.",
            desc=f"{len(libs)} bibliotecas analisadas. O script de permissões não exporta labels — é necessário verificar manualmente ou via PowerShell.\n\nComando para verificar:\nGet-PnPSite -Includes SensitivityLabel | Select SensitivityLabel",
            rec="ANTES DE ATIVAR COPILOT:\n• Configurar Sensitivity Labels no Microsoft Purview:\n  - 'Público' → Copilot pode usar livremente\n  - 'Interno' → Copilot usa mas não partilha externamente\n  - 'Confidencial' → Copilot restringido\n  - 'Altamente Confidencial' → Copilot bloqueado\n• Aplicar labels obrigatórios a TODAS as bibliotecas.\n• Ativar auto-labeling para documentos com dados sensíveis (PII, financeiros, etc.).\n• Usar DLP para forçar labels em documentos novos.",
            affected=[f"📚 {l} — verificar se tem sensitivity label" for l in sorted(libs)],
            metric=str(len(libs)), mlabel="Bibliotecas s/ label verificado", thr="0 sem label")

    # ---- CR-004: Herança quebrada ----
    def _cr004(self):
        r = self.stats["unique_ratio"]; u = self.stats["unique_count"]; t = self.stats["total"]
        sev = "critico" if r > 50 else "alto" if r > 30 else "medio" if r > 10 else "baixo" if u > 0 else "ok"
        paths = list(set(rec.get("caminho", rec.get("recurso", ""))
                        for rec in self.recs if is_uniq(rec.get("unicas", ""))))[:20]
        self._add(id="CR-004", title="Herança Quebrada Expõe Dados Imprevistos", cat="Estrutura", sev=sev,
            copilot_risk="Quando a herança é quebrada, utilizadores podem ter acesso a pastas profundas que desconhecem. O Copilot encontra TUDO — incluindo pastas 'escondidas' onde o utilizador nunca navegou mas tem acesso técnico.",
            scenario="Admin adicionou Maria a uma pasta há 2 anos. Maria nunca a visitou. Agora Maria pergunta ao Copilot: 'O que temos sobre o projeto X?' → Copilot surfacea documentos dessa pasta oculta.",
            desc=f"{u} de {t} atribuições ({r}%) com herança quebrada.",
            rec="• Fazer inventário de TODAS as quebras de herança.\n• Perguntar: 'Este utilizador SABE que tem acesso aqui?'\n• Se não → remover acesso ou reverter herança.\n• Reorganizar em bibliotecas separadas (CR-008).\n• Ativar RCD em sites sensíveis temporariamente.",
            affected=paths, metric=f"{r}%", mlabel="Herança quebrada", thr="< 10%")

    # ---- CR-005: Permissões diretas ----
    def _cr005(self):
        du = defaultdict(list)
        for r in self.recs:
            if is_user(r.get("tipo_membro", "")):
                du[r.get("membro", "?")].append(r.get("caminho", r.get("recurso", "")))
        c = len(du); ratio = self.stats["direct_user_ratio"]
        sev = "alto" if ratio > 40 else "medio" if ratio > 20 else "baixo" if c > 0 else "ok"
        aff = [f"{u} → acesso direto em {len(p)} locais" for u, p in sorted(du.items(), key=lambda x: -len(x[1]))][:15]
        self._add(id="CR-005", title="Permissões Diretas Dificultam Auditoria", cat="Governança", sev=sev,
            copilot_risk="Permissões diretas são invisíveis a nível de grupo. Quando um colaborador muda de função, a permissão continua. O Copilot acede a tudo o que o utilizador tecnicamente pode aceder — mesmo que não devesse.",
            scenario="João mudou do Financeiro para Marketing. Mantém acesso direto a pasta de contratos. Pergunta ao Copilot: 'Mostra-me contratos recentes' → Copilot mostra contratos que já não deveria ver.",
            desc=f"{c} utilizadores com permissões diretas ({ratio}%).",
            rec="• ANTES do Copilot: migrar TODAS as permissões diretas para grupos.\n• Criar grupos por função: 'GRP-Financeiro-Leitura', 'GRP-RH-Edição', etc.\n• Implementar Azure AD Access Reviews trimestrais.\n• Ao mudar de função → remover do grupo antigo, adicionar ao novo.",
            affected=aff, metric=str(c), mlabel="Utilizadores diretos", thr="0")

    # ---- CR-006: Oversharing ----
    def _cr006(self):
        path_mems = defaultdict(set)
        for r in self.recs:
            p = r.get("caminho", r.get("recurso", ""))
            if p and r.get("membro"): path_mems[p].add(r.get("membro", ""))
        overshared = [(p, len(ms)) for p, ms in path_mems.items() if len(ms) >= 10]
        overshared.sort(key=lambda x: -x[1])
        sev = "critico" if len(overshared) > 10 else "alto" if len(overshared) > 5 else "medio" if overshared else "ok"
        aff = [f"{p} → acessível por {c} membros" for p, c in overshared[:15]]
        self._add(id="CR-006", title="Locais com Oversharing (10+ Membros)", cat="Exposição de Dados", sev=sev,
            copilot_risk="Quanto mais pessoas têm acesso, mais o Copilot pode surfacear o conteúdo para utilizadores diferentes. Dados 'escondidos' em pastas partilhadas tornam-se encontráveis.",
            scenario="Pasta 'Gestão/Relatórios' partilhada com 25 pessoas. Contém relatório de performance individual. Qualquer um dos 25 pergunta ao Copilot e vê avaliações de colegas.",
            desc=f"{len(overshared)} locais acessíveis por 10+ membros cada.",
            rec="• Rever CADA local com 10+ membros: o acesso é necessário para todos?\n• Aplicar princípio do menor privilégio.\n• Separar conteúdo sensível em bibliotecas restritas.\n• Usar sensitivity labels para conteúdo confidencial dentro destas pastas.",
            affected=aff, metric=str(len(overshared)), mlabel="Locais overshared", thr="Mínimo")

    # ---- CR-007: Candidatos a RCD ----
    def _cr007(self):
        # Sites/bibliotecas com dados sensíveis (FC + herança quebrada + muitos membros)
        lib_risk = defaultdict(lambda: {"fc": 0, "unique": 0, "members": set(), "paths": set()})
        for r in self.recs:
            p = r.get("caminho", ""); lib = p.split("/")[0] if p else r.get("recurso", "")
            if is_fc(r.get("niveis", "")): lib_risk[lib]["fc"] += 1
            if is_uniq(r.get("unicas", "")): lib_risk[lib]["unique"] += 1
            lib_risk[lib]["members"].add(r.get("membro", ""))
            lib_risk[lib]["paths"].add(p)
        candidates = []
        for lib, info in lib_risk.items():
            risk_score = info["fc"] * 3 + info["unique"] * 2 + len(info["members"])
            if risk_score > 20 or info["fc"] > 3:
                candidates.append((lib, risk_score, info))
        candidates.sort(key=lambda x: -x[1])
        sev = "alto" if len(candidates) > 3 else "medio" if candidates else "ok"
        aff = [f"📚 {lib} — risco: {sc} (FC:{i['fc']}, únicos:{i['unique']}, membros:{len(i['members'])})"
               for lib, sc, i in candidates[:10]]
        self._add(id="CR-007", title="Sites Candidatos a Restricted Content Discovery", cat="Copilot Controls", sev=sev,
            copilot_risk="O Restricted Content Discovery (RCD) impede o Copilot de descobrir conteúdo de sites específicos. Sites com dados sensíveis e muitas permissões devem ter RCD ativo ANTES de ativar Copilot.",
            scenario="Site de RH sem RCD. Qualquer utilizador pergunta: 'Mostra-me informações sobre demissões recentes' → Copilot encontra e apresenta os dados.",
            desc=f"{len(candidates)} bibliotecas/sites com perfil de risco elevado para Copilot.",
            rec="ATIVAR RCD nos sites identificados:\n\nPowerShell:\nSet-SPOSite -Identity <site-url> -RestrictContentOrgWideSearch $true\n\n• RCD impede o Copilot de descobrir o conteúdo (não altera permissões).\n• Utilizadores com acesso direto ainda podem navegar manualmente.\n• Usar RCD como medida temporária enquanto se reorganizam permissões.\n• Para bloqueio total: considerar Restricted SharePoint Search (RSS).",
            affected=aff, metric=str(len(candidates)), mlabel="Sites p/ RCD", thr="Avaliar cada")

    # ---- CR-008: Divisão de bibliotecas ----
    def _cr008(self):
        lfd = self._lib_folder_data()
        lib_mems = defaultdict(set)
        for r in self.recs:
            if r.get("_sk") == "bibliotecas": lib_mems[r.get("recurso", "")].add(r.get("membro", ""))
        recs = []
        for lib, records in lfd.items():
            flv = defaultdict(lambda: {"mems": set(), "cnt": 0})
            for r in records:
                parts = r.get("caminho", "").split("/")
                if len(parts) >= 2:
                    flv[parts[1]]["mems"].add(r.get("membro", "")); flv[parts[1]]["cnt"] += 1
            ufl = {k: v for k, v in flv.items() if v["cnt"] > 0}
            if len(ufl) >= 3:
                pm = lib_mems.get(lib, set())
                cands = []
                for f, info in ufl.items():
                    ov = len(info["mems"] & pm) / max(len(pm), 1) if pm else 0
                    cands.append(f"  📁 {lib}/{f} → {len(info['mems'])} membros, {round(ov*100)}% sobreposição")
                recs.append({"lib": lib, "folders": len(ufl), "cands": cands})
        sev = "alto" if len(recs) > 2 else "medio" if recs else "ok"
        aff = []
        for rc in recs:
            aff.append(f"📚 {rc['lib']} — {rc['folders']} pastas com público diferente:")
            aff.extend(rc["cands"][:5])
        self._add(id="CR-008", title="Bibliotecas a Dividir para Isolar Dados", cat="Arquitetura", sev=sev,
            copilot_risk="Quando dados de diferentes equipas/departamentos estão na mesma biblioteca, o Copilot pode cruzar informações. Dividir em bibliotecas isoladas cria fronteiras naturais de segurança.",
            scenario="Biblioteca 'Documentos' contém pasta RH e pasta Financeiro. Colaborador do Financeiro pergunta: 'Quais documentos tenho acesso?' → Copilot mostra documentos de RH se a herança foi quebrada incorretamente.",
            desc=f"{len(recs)} biblioteca(s) com pastas de público-alvo diferente.",
            rec="• Criar biblioteca separada por departamento/equipa.\n• Mover conteúdo e aplicar permissões ao nível da biblioteca.\n• Elimina necessidade de quebrar herança em pastas.\n• Usar sensitivity labels diferentes por biblioteca.\n• Exemplo: 'Documentos' → 'Bibl. RH' + 'Bibl. Financeiro' + 'Bibl. Marketing'.",
            affected=aff, metric=str(len(recs)), mlabel="Bibl. a dividir", thr="0")

    # ---- CR-009: Edição no site raiz ----
    def _cr009(self):
        ed = []
        for r in self.recs:
            if r.get("_sk") == "site":
                n = r.get("niveis", "")
                if is_edit(n) or is_fc(n):
                    ed.append(f"{r.get('membro', '')} → {n}")
        sev = "alto" if len(ed) > 5 else "medio" if ed else "ok"
        self._add(id="CR-009", title="Edição no Site Raiz Propaga para Copilot", cat="Herança", sev=sev,
            copilot_risk="Permissões de edição ao nível do site propagam para TODAS as bibliotecas e pastas por herança. O Copilot vê este utilizador como tendo acesso a tudo.",
            scenario="Ana tem 'Contribute' no site raiz. Nunca acedeu à biblioteca de Jurídico. Pergunta ao Copilot: 'Mostra-me contratos' → Copilot encontra contratos jurídicos porque Ana tem acesso técnico via herança.",
            desc=f"{len(ed)} membros com edição+ no site raiz.",
            rec="• Remover permissões de edição do site raiz.\n• Atribuir 'Read' ao nível do site.\n• Conceder 'Edit'/'Contribute' APENAS nas bibliotecas específicas.\n• Isto limita automaticamente o que o Copilot pode surfacear por utilizador.",
            affected=ed[:15], metric=str(len(ed)), mlabel="Editores no raiz", thr="Mínimo")

    # ---- CR-010: Níveis que permitem partilha ----
    def _cr010(self):
        share_levels = []
        for r in self.recs:
            n = r.get("niveis", "")
            if is_fc(n) or is_edit(n):
                share_levels.append(r)
        # % de atribuições que podem partilhar
        share_pct = round(len(share_levels) / max(self.stats["total"], 1) * 100, 1)
        sev = "alto" if share_pct > 50 else "medio" if share_pct > 30 else "baixo" if share_levels else "ok"
        self._add(id="CR-010", title="Permissões que Permitem Partilha Via Copilot", cat="Partilha", sev=sev,
            copilot_risk="Utilizadores com Edit/Contribute/FC podem criar links de partilha. O Copilot pode sugerir 'partilhar este documento' e o utilizador expande o acesso inadvertidamente.",
            scenario="Copilot sugere: 'Quer partilhar este relatório com a equipa?' → utilizador clica sim → documento sensível fica acessível a mais pessoas → Copilot deles também passa a aceder.",
            desc=f"{len(share_levels)} atribuições ({share_pct}%) permitem partilha de conteúdo.",
            rec="• Configurar políticas de partilha no SharePoint Admin:\n  - Desativar 'Anyone' links.\n  - Limitar a 'People in your organization' ou 'Specific people'.\n  - Definir expiração de links partilhados.\n• Configurar DLP no Purview para bloquear partilha de documentos com labels sensíveis.\n• Ativar 'Require approval' para partilha em sites sensíveis.",
            affected=[f"{r.get('membro', '')} → {r.get('niveis', '')} em {r.get('caminho', r.get('recurso', ''))}"
                      for r in share_levels[:10]],
            metric=f"{share_pct}%", mlabel="Podem partilhar", thr="Mínimo necessário")

    # ---- CR-011: Órfãs ----
    def _cr011(self):
        sus = []
        for r in self.recs:
            lo = r.get("login", "").lower(); me = r.get("membro", "").lower()
            p = r.get("caminho", r.get("recurso", ""))
            if re.search(r"s-1-5-\d+-\d+", lo): sus.append(f"SID: {r.get('login', '')} → {p}")
            elif any(x in me for x in ["removed", "deleted", "removido", "eliminado"]): sus.append(f"Removida: {r.get('membro', '')} → {p}")
        sev = "alto" if len(sus) > 3 else "medio" if sus else "ok"
        self._add(id="CR-011", title="Contas Órfãs com Acesso Ativo", cat="Segurança", sev=sev,
            copilot_risk="Contas órfãs (SIDs não resolvidos, eliminadas) podem ser reutilizadas ou indicam acesso fantasma. Se associadas a utilizadores ativos por coincidência de permissão, o Copilot pode surfacear dados que não deveriam existir.",
            scenario="SID antigo aponta para acesso a pasta de Direção. Novo colaborador herda indiretamente. Copilot surfacea atas de direção.",
            desc=f"{len(sus)} permissões órfãs encontradas.",
            rec="• Remover IMEDIATAMENTE todas as permissões com SIDs.\n• Remover contas eliminadas/removidas.\n• Executar cleanup ANTES de ativar Copilot.\n• Implementar Azure AD Access Reviews automáticas.",
            affected=sus[:15], metric=str(len(sus)), mlabel="Órfãs", thr="0")

    # ---- CR-012: Inconsistência ----
    def _cr012(self):
        ml = defaultdict(set)
        for r in self.recs:
            m = r.get("membro", ""); n = r.get("niveis", "")
            if m and n: ml[m].add(n)
        inc = [(m, sorted(l)) for m, l in ml.items() if len(l) > 2]
        sev = "medio" if len(inc) > 5 else "baixo" if inc else "ok"
        aff = [f"{m}: {', '.join(ls)}" for m, ls in inc[:15]]
        self._add(id="CR-012", title="Inconsistência de Acesso (Dados Imprevistos)", cat="Governança", sev=sev,
            copilot_risk="Utilizadores com níveis inconsistentes (Read num local, FC noutro) têm comportamento imprevisível com o Copilot. O Copilot pode misturar dados de diferentes níveis de sensibilidade nas respostas.",
            scenario="Maria tem Read em Marketing e Full Control em Financeiro. Pergunta: 'Mostra-me relatórios' → Copilot mistura relatórios de marketing (ok) com relatórios financeiros confidenciais.",
            desc=f"{len(inc)} membros com 3+ níveis de permissão diferentes.",
            rec="• Padronizar perfis: cada utilizador deve ter nível consistente.\n• Criar grupos com nível fixo.\n• Remover permissões acumuladas.\n• Usar Azure AD Access Reviews para limpar.",
            affected=aff, metric=str(len(inc)), mlabel="Inconsistentes", thr="0 ideal")

    # ---- CR-013: Profundidade ----
    def _cr013(self):
        dp = set()
        for r in self.recs:
            if is_uniq(r.get("unicas", "")) and r.get("_sk") == "subpastas":
                p = r.get("caminho", ""); d = p.count("/")
                if d >= 3: dp.add((p, d))
        dp = sorted(dp, key=lambda x: -x[1])
        sev = "alto" if len(dp) > 5 else "medio" if dp else "ok"
        self._add(id="CR-013", title="Pastas Profundas Que o Copilot Descobrirá", cat="Descoberta", sev=sev,
            copilot_risk="Pastas profundas com permissões únicas são frequentemente 'escondidas' — utilizadores nem sabem que existem. Mas o Copilot pesquisa TODO o conteúdo acessível, incluindo pastas onde o utilizador nunca navegou.",
            scenario="Pasta em Documentos/Projetos/2023/Arquivados/Confidencial com herança quebrada. Utilizador nunca viu esta pasta. Copilot: 'Encontrei informação relevante em Documentos/Projetos/2023/Arquivados/Confidencial...'",
            desc=f"{len(dp)} pastas a 3+ níveis de profundidade com permissões únicas.",
            rec="• Mover conteúdo sensível para bibliotecas restritas.\n• Eliminar pastas profundas desnecessárias.\n• Ativar RCD em sites com estruturas profundas.\n• Simplificar: máximo 2 níveis de pastas.",
            affected=[f"[Nível {d}] {p}" for p, d in dp[:15]],
            metric=str(len(dp)), mlabel="Pastas profundas", thr="0")

    # ---- CR-014: Mistura de dados ----
    def _cr014(self):
        # Analisar bibliotecas onde existem membros com FC e membros com Read
        lib_levels = defaultdict(lambda: {"fc": set(), "edit": set(), "read": set()})
        for r in self.recs:
            if r.get("_sk") in ("bibliotecas", "subpastas"):
                p = r.get("caminho", ""); lib = p.split("/")[0] if p else r.get("recurso", "")
                m = r.get("membro", ""); n = r.get("niveis", "")
                if is_fc(n): lib_levels[lib]["fc"].add(m)
                elif is_edit(n): lib_levels[lib]["edit"].add(m)
                elif is_read(n): lib_levels[lib]["read"].add(m)
        mixed = []
        for lib, levels in lib_levels.items():
            if (levels["fc"] or levels["edit"]) and levels["read"]:
                mixed.append(f"📚 {lib}: {len(levels['fc'])} FC + {len(levels['edit'])} Edit + {len(levels['read'])} Read")
        sev = "medio" if len(mixed) > 3 else "baixo" if mixed else "ok"
        self._add(id="CR-014", title="Mistura de Níveis de Acesso na Mesma Biblioteca", cat="Classificação", sev=sev,
            copilot_risk="Quando a mesma biblioteca tem utilizadores com diferentes níveis, o Copilot pode gerar respostas que revelam a EXISTÊNCIA de documentos que o utilizador de Read não deveria saber que existem (mesmo sem mostrar o conteúdo).",
            scenario="Biblioteca tem utilizadores Read e FC. Utilizador Read pergunta: 'Que tipos de documentos temos?' → Copilot revela a existência de pastas que o utilizador Read não deveria conhecer.",
            desc=f"{len(mixed)} bibliotecas com mistura de níveis de acesso.",
            rec="• Separar conteúdo por nível de acesso em bibliotecas diferentes.\n• Conteúdo sensível → biblioteca 'Confidencial' com menos membros.\n• Conteúdo geral → biblioteca 'Público' com mais membros.\n• Usar sensitivity labels para forçar a separação.",
            affected=mixed[:10], metric=str(len(mixed)), mlabel="Bibl. misturadas", thr="Mínimo")

    # ---- CR-015: Checklist ----
    def _cr015(self):
        checks = []
        s = self.stats
        # 1. Grupos amplos
        f001 = next((f for f in self.findings if f["id"] == "CR-001"), None)
        checks.append(("Remover grupos 'Everyone/Todos'", f001 and f001["sev"] == "ok",
                        "Nenhum grupo amplo" if f001 and f001["sev"] == "ok" else "AÇÃO NECESSÁRIA"))
        # 2. Permissões diretas
        checks.append(("Migrar permissões diretas para grupos", s["direct_users"] == 0,
                        "0 diretos" if s["direct_users"] == 0 else f"{s['direct_users']} a migrar"))
        # 3. Full Control
        checks.append(("Reduzir Full Control para < 5%", s["fc_ratio"] <= 5,
                        f"{s['fc_ratio']}%" if s["fc_ratio"] <= 5 else f"{s['fc_ratio']}% — reduzir"))
        # 4. Herança
        checks.append(("Herança quebrada < 10%", s["unique_ratio"] <= 10,
                        f"{s['unique_ratio']}%" if s["unique_ratio"] <= 10 else f"{s['unique_ratio']}% — simplificar"))
        # 5. Dividir bibliotecas
        checks.append(("Dividir bibliotecas com públicos mistos", s["libs_split"] == 0,
                        "Nenhuma" if s["libs_split"] == 0 else f"{s['libs_split']} a dividir"))
        # 6. Labels (sempre pendente - precisa verificação manual)
        checks.append(("Configurar Sensitivity Labels (Purview)", False,
                        "Verificação manual necessária"))
        # 7. RCD
        checks.append(("Ativar RCD em sites sensíveis", False,
                        "Verificação manual necessária"))
        # 8. Órfãs
        f011 = next((f for f in self.findings if f["id"] == "CR-011"), None)
        orphans = f011 and f011["sev"] == "ok" if f011 else True
        checks.append(("Limpar permissões órfãs", orphans,
                        "Limpo" if orphans else "LIMPEZA NECESSÁRIA"))
        # 9. DLP
        checks.append(("Configurar DLP no Purview", False, "Verificação manual necessária"))
        # 10. Partilha
        checks.append(("Restringir políticas de partilha", False, "Verificação manual necessária"))
        # 11. Access Reviews
        checks.append(("Ativar Azure AD Access Reviews", False, "Verificação manual necessária"))
        # 12. Treino
        checks.append(("Treinar utilizadores sobre Copilot e dados", False, "Verificação manual necessária"))

        done = sum(1 for _, ok, _ in checks if ok)
        total_checks = len(checks)
        pct = round(done / total_checks * 100)
        sev = "ok" if pct >= 80 else "baixo" if pct >= 60 else "medio" if pct >= 40 else "alto" if pct >= 20 else "critico"

        aff = []
        for label, ok, detail in checks:
            icon = "✅" if ok else "❌"
            aff.append(f"{icon} {label} — {detail}")

        self._add(id="CR-015", title="Checklist Copilot-Ready", cat="Governança", sev=sev,
            copilot_risk="Esta checklist resume TODAS as ações necessárias antes de ativar o Microsoft 365 Copilot. Cada item não cumprido representa um vetor de vazamento de dados.",
            scenario="Organização ativa Copilot sem preparação → utilizadores perguntam sobre salários, contratos, avaliações → Copilot responde com dados que não deveriam ser acessíveis.",
            desc=f"Checklist: {done}/{total_checks} itens cumpridos ({pct}%).",
            rec="Cumprir TODOS os itens antes de ativar Copilot.\n\nOrdem recomendada:\n1. Remover grupos amplos (CR-001)\n2. Limpar órfãs (CR-011)\n3. Reduzir Full Control (CR-002)\n4. Migrar permissões diretas (CR-005)\n5. Dividir bibliotecas (CR-008)\n6. Configurar Sensitivity Labels\n7. Ativar RCD em sites sensíveis\n8. Configurar DLP\n9. Restringir partilha\n10. Ativar Access Reviews\n11. Reduzir herança quebrada (CR-004)\n12. Treinar utilizadores",
            affected=aff, metric=f"{pct}%", mlabel="Copilot-Ready", thr="100%")


# ============================================================
# Gráficos
# ============================================================
def chart_score(score):
    if not HAS_MPL: return None
    fig, ax = plt.subplots(figsize=(3, 3))
    c = "#ef4444" if score < 40 else "#f59e0b" if score < 70 else "#10b981"
    ax.pie([score, 100-score], colors=[c, "#334155"], startangle=90, wedgeprops=dict(width=0.35, edgecolor="none"))
    ax.text(0, 0.05, f"{score}", ha="center", va="center", fontsize=28, fontweight="bold", color=c)
    ax.text(0, -0.18, "/ 100", ha="center", va="center", fontsize=9, color="#94a3b8")
    ax.text(0, -0.42, "Copilot-Ready", ha="center", va="center", fontsize=8, color="#64748b")
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=150, transparent=True, bbox_inches="tight")
    plt.close(fig); buf.seek(0); return buf

def chart_risk(findings):
    if not HAS_MPL: return None
    sc = Counter(f["sev"] for f in findings)
    labels = ["Crítico", "Alto", "Médio", "Baixo", "OK"]
    keys = ["critico", "alto", "medio", "baixo", "ok"]
    vals = [sc.get(k, 0) for k in keys]
    cols = ["#ef4444", "#f97316", "#f59e0b", "#3b82f6", "#10b981"]
    fig, ax = plt.subplots(figsize=(5, 2.5))
    bars = ax.barh(labels, vals, color=cols, height=0.6, edgecolor="none")
    ax.set_xlim(0, max(vals)+1 if max(vals)>0 else 2); ax.invert_yaxis()
    for b, v in zip(bars, vals): ax.text(b.get_width()+0.15, b.get_y()+b.get_height()/2, str(v), va="center", fontsize=10, fontweight="bold", color="#333")
    ax.spines["top"].set_visible(False); ax.spines["right"].set_visible(False); ax.spines["bottom"].set_visible(False)
    ax.tick_params(bottom=False, labelbottom=False, axis="y", labelsize=10)
    ax.set_title("Distribuição de Riscos Copilot", fontsize=11, fontweight="bold", pad=10)
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=150, bbox_inches="tight"); plt.close(fig); buf.seek(0); return buf


# ============================================================
# DOCX helpers
# ============================================================
def shade(cell, color):
    cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>'))

def cell_txt(cell, text, bold=False, size=9, color=None, align=None):
    cell.text = ""; p = cell.paragraphs[0]
    if align: p.alignment = align
    r = p.add_run(str(text)); r.font.size = Pt(size); r.font.bold = bold
    if color: r.font.color.rgb = color
    p.paragraph_format.space_before = Pt(2); p.paragraph_format.space_after = Pt(2)

def add_table(doc, headers, rows, widths=None, hdr_color="1E3A5F"):
    t = doc.add_table(rows=1+len(rows), cols=len(headers)); t.alignment = WD_TABLE_ALIGNMENT.CENTER; t.style = "Table Grid"
    for i, h in enumerate(headers): shade(t.rows[0].cells[i], hdr_color); cell_txt(t.rows[0].cells[i], h, True, 9, C_WHITE)
    for ri, row in enumerate(rows):
        for ci, v in enumerate(row):
            if ci < len(headers):
                cell_txt(t.rows[ri+1].cells[ci], str(v), size=8)
                if ri % 2 == 1: shade(t.rows[ri+1].cells[ci], "F8FAFC")
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                if i < len(row.cells): row.cells[i].width = Cm(w)
    return t

def heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    for r in h.runs: r.font.color.rgb = C_BLUE
    return h

def para(doc, text, bold=False, size=10, color=None, after=6):
    p = doc.add_paragraph(); r = p.add_run(text); r.font.size = Pt(size); r.font.bold = bold
    if color: r.font.color.rgb = color
    p.paragraph_format.space_after = Pt(after); return p


# ============================================================
# Gerador DOCX principal
# ============================================================
def generate_docx(findings, stats, data, excel_file, output):
    doc = Document()
    res = data["resumo"]
    site = res.get("Site", res.get("site", "SharePoint"))
    url = res.get("URL", res.get("url", ""))
    now = datetime.now(); score = stats["copilot_score"]
    style = doc.styles["Normal"]; style.font.name = "Calibri"; style.font.size = Pt(10); style.font.color.rgb = C_TEXT
    for s in doc.sections: s.top_margin = Cm(2); s.bottom_margin = Cm(2); s.left_margin = Cm(2.5); s.right_margin = Cm(2.5)

    # ---- CAPA ----
    for _ in range(3): doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("🤖 Preparação para Microsoft 365 Copilot"); r.font.size = Pt(26); r.font.bold = True; r.font.color.rgb = C_BLUE
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Relatório de Riscos de Vazamento de Dados"); r.font.size = Pt(16); r.font.color.rgb = C_TEXT2
    doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(site); r.font.size = Pt(14); r.font.bold = True
    if url:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(url); r.font.size = Pt(10); r.font.color.rgb = C_BLUE2
    for _ in range(4): doc.add_paragraph()
    mt = doc.add_table(rows=5, cols=2); mt.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, (l, v) in enumerate([("Data", now.strftime("%d/%m/%Y %H:%M")), ("Fonte", os.path.basename(excel_file)),
        ("Atribuições Analisadas", str(stats["total"])), ("Score Copilot-Ready", f"{score}/100"),
        ("Risco de Vazamento", "Elevado" if score < 40 else "Moderado" if score < 70 else "Baixo")]):
        cell_txt(mt.rows[i].cells[0], l, True, 10, C_TEXT2); cell_txt(mt.rows[i].cells[1], v, size=10)
    doc.add_page_break()

    # ---- ÍNDICE ----
    heading(doc, "Índice", 1)
    toc = ["1. Por Que Este Relatório É Crítico", "2. Como o Copilot Acede aos Dados",
           "3. Score Copilot-Ready", "4. Métricas de Risco", "5. Análise de Riscos (15 verificações)"]
    for i, f in enumerate(findings): toc.append(f"   5.{i+1}  {f['id']} — {f['title']}")
    toc += ["6. Checklist de Ações", "7. Plano de Implementação", "8. Configurações Técnicas", "9. Referências"]
    for item in toc:
        p = doc.add_paragraph(item); p.paragraph_format.space_after = Pt(2)
        for r in p.runs: r.font.size = Pt(10); r.font.color.rgb = C_TEXT if not item.startswith("   ") else C_TEXT2
    doc.add_page_break()

    # ---- 1. POR QUE ----
    heading(doc, "1. Por Que Este Relatório É Crítico", 1)
    para(doc, "O Microsoft 365 Copilot é um assistente de IA que acede a TODOS os dados que o utilizador tem permissão para ver no SharePoint, OneDrive, Teams, Exchange e mais.", size=10)
    para(doc, "⚠️ O PROBLEMA:", bold=True, size=11, color=C_RED)
    bullets = [
        "O Copilot NÃO valida se o utilizador DEVERIA ver os dados — apenas se TEM acesso técnico.",
        "Dados 'escondidos' em pastas profundas ou bibliotecas esquecidas tornam-se encontráveis.",
        "Permissões excessivas que hoje passam despercebidas serão ATIVAMENTE exploradas pelo Copilot.",
        "Uma simples pergunta como 'Quais são os salários da equipa?' pode expor dados confidenciais.",
        "O Copilot pode cruzar dados de diferentes fontes, criando fugas de informação compostas.",
    ]
    for b in bullets:
        p = doc.add_paragraph(b, style="List Bullet"); p.paragraph_format.space_after = Pt(3)
    para(doc, "Este relatório identifica EXATAMENTE onde estão os riscos no seu SharePoint e o que fazer antes de ativar o Copilot.", bold=True, size=10, color=C_BLUE)
    doc.add_page_break()

    # ---- 2. COMO O COPILOT ACEDE ----
    heading(doc, "2. Como o Microsoft 365 Copilot Acede aos Dados", 1)
    flow = [
        ("1. Pergunta do Utilizador", "O utilizador faz uma pergunta ao Copilot (ex: 'Mostra-me os relatórios recentes')."),
        ("2. Pesquisa em Todo o M365", "O Copilot pesquisa SharePoint, OneDrive, Teams, Exchange — TUDO acessível ao utilizador."),
        ("3. Filtragem por Permissões", "O Copilot aplica as permissões do utilizador. Se o utilizador TEM acesso, o Copilot VÊ."),
        ("4. Geração de Resposta", "O Copilot combina informações de múltiplas fontes numa resposta unificada."),
        ("5. O RISCO", "Se as permissões estão mal configuradas, o Copilot expõe dados que o utilizador não deveria ver."),
    ]
    for title, desc in flow:
        para(doc, title, bold=True, size=10, color=C_BLUE, after=2)
        para(doc, desc, size=9, after=8)

    para(doc, "Controlos disponíveis:", bold=True, size=11, color=C_PURPLE)
    controls = [
        ("Permissões SharePoint", "Base — o Copilot respeita estas permissões."),
        ("Sensitivity Labels (Purview)", "Classificam e protegem conteúdo. O Copilot respeita labels."),
        ("Restricted Content Discovery (RCD)", "Impede o Copilot de DESCOBRIR conteúdo de sites específicos."),
        ("Restricted SharePoint Search (RSS)", "Limita o Copilot a uma lista curada de sites (máx. 100)."),
        ("DLP (Data Loss Prevention)", "Bloqueia partilha de conteúdo sensível."),
        ("Azure AD Access Reviews", "Recertificação periódica de acessos."),
    ]
    add_table(doc, ["Controlo", "Descrição"], [[c, d] for c, d in controls], [5, 10], "6B21A8")
    doc.add_page_break()

    # ---- 3. SCORE ----
    heading(doc, "3. Score Copilot-Ready", 1)
    sc = chart_score(score)
    if sc:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(sc, width=Inches(2.2))
    label = "Pronto" if score >= 80 else "Quase Pronto" if score >= 60 else "Risco Moderado" if score >= 40 else "NÃO ATIVAR COPILOT"
    col = C_GREEN if score >= 80 else C_YELLOW if score >= 60 else C_ORANGE if score >= 40 else C_RED
    para(doc, f"{label}", bold=True, size=14, color=col)
    if score < 40:
        para(doc, "⛔ RECOMENDAÇÃO: NÃO ativar o Microsoft 365 Copilot até resolver os problemas identificados neste relatório. O risco de vazamento de dados é elevado.", size=11, color=C_RED)
    elif score < 70:
        para(doc, "⚠️ RECOMENDAÇÃO: Resolver os itens CRÍTICOS e ALTOS antes de ativar. Ativar com Restricted SharePoint Search (RSS) como medida temporária.", size=11, color=C_ORANGE)
    else:
        para(doc, "✅ O ambiente está razoavelmente preparado. Recomenda-se ativar com monitorização e resolver os itens restantes.", size=11, color=C_GREEN)

    sev_chart = chart_risk(findings)
    if sev_chart:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(sev_chart, width=Inches(4))
    doc.add_page_break()

    # ---- 4. MÉTRICAS ----
    heading(doc, "4. Métricas de Risco para Copilot", 1)
    s = stats
    metrics = [
        ["Métrica", "Valor", "Limite Copilot", "Risco"],
        ["Permissões únicas", f"{s['unique_ratio']}%", "< 10%", "🔴" if s["unique_ratio"] > 30 else "🟠" if s["unique_ratio"] > 10 else "✅"],
        ["Utilizadores diretos", str(s["direct_users"]), "0", "🔴" if s["direct_users"] > 10 else "🟠" if s["direct_users"] > 0 else "✅"],
        ["Full Control", f"{s['fc_ratio']}%", "< 5%", "🔴" if s["fc_ratio"] > 20 else "🟠" if s["fc_ratio"] > 5 else "✅"],
        ["Membros com acesso amplo", str(s["wide_access_members"]), "Mínimo", "🟠" if s["wide_access_members"] > 5 else "✅"],
        ["Locais overshared", str(s["overshared_paths"]), "0", "🔴" if s["overshared_paths"] > 5 else "🟠" if s["overshared_paths"] > 0 else "✅"],
        ["Bibliotecas a dividir", str(s["libs_split"]), "0", "🟠" if s["libs_split"] > 0 else "✅"],
        ["Total membros", str(s["total_members"]), "—", "ℹ️"],
        ["Bibliotecas", str(s["lib_count"]), "—", "ℹ️"],
    ]
    add_table(doc, metrics[0], metrics[1:], [5, 3, 3, 2])
    doc.add_page_break()

    # ---- 5. ANÁLISE DE RISCOS ----
    heading(doc, "5. Análise de Riscos para Copilot", 1)
    para(doc, "Cada risco inclui: descrição do impacto no Copilot, cenário real de vazamento, e ação corretiva.", size=10)

    for idx, f in enumerate(findings):
        sev = f["sev"]; sl = SEV_L.get(sev, sev); si = SEV_I.get(sev, ""); sc = SEV_C.get(sev, C_TEXT)
        h = doc.add_heading(f"5.{idx+1}  {f['id']} — {f['title']}", level=2)
        for r in h.runs: r.font.color.rgb = C_BLUE; r.font.size = Pt(12)

        # Severidade
        p = doc.add_paragraph()
        r = p.add_run(f"{si} {sl}"); r.font.bold = True; r.font.size = Pt(11); r.font.color.rgb = sc
        r = p.add_run(f"  |  {f['metric']} {f['mlabel']}  |  {f['cat']}  |  Meta: {f['thr']}")
        r.font.size = Pt(9); r.font.color.rgb = C_TEXT2

        # Risco Copilot
        para(doc, "🤖 Risco com o Copilot:", bold=True, size=10, color=C_RED, after=2)
        para(doc, f["copilot_risk"], size=9, after=6)

        # Cenário
        para(doc, "💬 Cenário de Vazamento:", bold=True, size=10, color=C_ORANGE, after=2)
        para(doc, f["scenario"], size=9, after=6)

        # Diagnóstico
        para(doc, "📋 Diagnóstico:", bold=True, size=10, color=C_BLUE, after=2)
        para(doc, f["desc"], size=9, after=6)

        # Recomendação
        para(doc, "✅ Ação Corretiva:", bold=True, size=10, color=C_GREEN, after=2)
        for line in f["rec"].split("\n"):
            line = line.strip()
            if line:
                p = doc.add_paragraph()
                r = p.add_run(line.lstrip("• ")); r.font.size = Pt(9)
                p.paragraph_format.space_after = Pt(2)

        # Afetados
        aff = f.get("affected", [])
        if aff:
            para(doc, f"📍 Itens Afetados ({len(aff)}):", bold=True, size=10, color=C_RED, after=2)
            at = doc.add_table(rows=min(len(aff), 15)+1, cols=1); at.style = "Table Grid"
            shade(at.rows[0].cells[0], "FEF2F2"); cell_txt(at.rows[0].cells[0], "Item", True, 8, C_RED)
            for ai, item in enumerate(aff[:15]):
                cell_txt(at.rows[ai+1].cells[0], str(item), size=8)
                if ai % 2 == 1: shade(at.rows[ai+1].cells[0], "FAFAFA")
        doc.add_paragraph()

    doc.add_page_break()

    # ---- 6. CHECKLIST ----
    heading(doc, "6. Checklist de Ações Antes de Ativar Copilot", 1)
    cr015 = next((f for f in findings if f["id"] == "CR-015"), None)
    if cr015 and cr015.get("affected"):
        rows_ck = []
        for item in cr015["affected"]:
            icon = "✅" if item.startswith("✅") else "❌"
            txt = item[2:].strip()
            rows_ck.append([icon, txt])
        add_table(doc, ["Estado", "Ação"], rows_ck, [2, 13], "6B21A8")
    doc.add_page_break()

    # ---- 7. PLANO ----
    heading(doc, "7. Plano de Implementação", 1)
    phases = [
        ("Fase 1 — IMEDIATA (Semana 1-2)", "Antes de qualquer ativação do Copilot:", [
            "Remover TODOS os grupos 'Everyone/Todos' (CR-001)",
            "Limpar permissões órfãs (CR-011)",
            "Ativar RCD nos sites sensíveis (CR-007)",
            "Reduzir Full Control (CR-002)"]),
        ("Fase 2 — CURTO PRAZO (Semana 3-4)", "Preparação estrutural:", [
            "Migrar permissões diretas para grupos (CR-005)",
            "Dividir bibliotecas com públicos mistos (CR-008)",
            "Configurar Sensitivity Labels no Purview (CR-003)",
            "Restringir políticas de partilha (CR-010)"]),
        ("Fase 3 — MÉDIO PRAZO (Mês 2)", "Activação controlada:", [
            "Ativar Copilot com Restricted SharePoint Search (RSS) — apenas sites aprovados",
            "Configurar DLP para bloquear partilha de conteúdo labelled como Confidencial",
            "Ativar Azure AD Access Reviews trimestrais"]),
        ("Fase 4 — CONTÍNUO", "Monitorização e melhoria:", [
            "Abrir gradualmente mais sites no RSS",
            "Monitorizar logs de acesso do Copilot no Purview Audit",
            "Revisões trimestrais de permissões",
            "Treinar utilizadores sobre uso seguro do Copilot"]),
    ]
    for title, desc, items in phases:
        para(doc, title, bold=True, size=12, color=C_BLUE, after=2)
        para(doc, desc, size=10, color=C_TEXT2, after=4)
        for item in items:
            p = doc.add_paragraph(item, style="List Bullet"); p.paragraph_format.space_after = Pt(2)
        doc.add_paragraph()
    doc.add_page_break()

    # ---- 8. CONFIGURAÇÕES TÉCNICAS ----
    heading(doc, "8. Configurações Técnicas (PowerShell)", 1)
    commands = [
        ("Ativar RCD num site", "Set-SPOSite -Identity https://contoso.sharepoint.com/sites/RH `\n  -RestrictContentOrgWideSearch $true"),
        ("Verificar RCD", "Get-SPOSite -Identity https://contoso.sharepoint.com/sites/RH `\n  | Select RestrictContentOrgWideSearch"),
        ("Ativar RSS (Restricted SharePoint Search)", "# No SharePoint Admin Center → Settings → Search\n# Ou via PowerShell:\nSet-SPOTenant -RestrictedSharePointSearch Enabled"),
        ("Adicionar site permitido ao RSS", "Add-SPORestrictedSearchAllowedList `\n  -SiteUrl https://contoso.sharepoint.com/sites/Aprovado"),
        ("Verificar Sensitivity Label do site", "Get-PnPSite -Includes SensitivityLabel | Select SensitivityLabel"),
        ("Listar permissões únicas", "Get-PnPList | Where {$_.HasUniqueRoleAssignments} | Select Title"),
    ]
    for title, cmd in commands:
        para(doc, title, bold=True, size=10, color=C_BLUE, after=2)
        p = doc.add_paragraph()
        r = p.add_run(cmd); r.font.size = Pt(8); r.font.name = "Consolas"
        r.font.color.rgb = C_TEXT2
        p.paragraph_format.space_after = Pt(10)
    doc.add_page_break()

    # ---- 9. REFERÊNCIAS ----
    heading(doc, "9. Referências", 1)
    refs = [
        ("Microsoft — Copilot Security & Privacy", "https://learn.microsoft.com/en-us/microsoft-365-copilot/microsoft-365-copilot-privacy"),
        ("Microsoft — Restricted Content Discovery", "https://learn.microsoft.com/en-us/sharepoint/restricted-content-discovery"),
        ("Microsoft — Restricted SharePoint Search", "https://learn.microsoft.com/en-us/sharepoint/restricted-sharepoint-search"),
        ("Microsoft — Sensitivity Labels", "https://learn.microsoft.com/en-us/purview/sensitivity-labels"),
        ("Microsoft — SharePoint Permissions Best Practices", "https://learn.microsoft.com/en-us/sharepoint/sites/best-practices-for-using-fine-grained-permissions"),
        ("Microsoft — DLP Policies", "https://learn.microsoft.com/en-us/purview/dlp-learn-about-dlp"),
        ("Microsoft — Azure AD Access Reviews", "https://learn.microsoft.com/en-us/entra/id-governance/access-reviews-overview"),
        ("Microsoft — SharePoint Advanced Management", "https://learn.microsoft.com/en-us/sharepoint/advanced-management"),
    ]
    for t, u in refs:
        para(doc, f"• {t}", bold=True, size=10, after=1)
        p = doc.add_paragraph(); r = p.add_run(f"  {u}"); r.font.size = Pt(8); r.font.color.rgb = C_BLUE2
        p.paragraph_format.space_after = Pt(6)

    para(doc, "Este relatório foi gerado automaticamente. Validar com administrador SharePoint e equipa de segurança antes de implementar.", size=8, color=C_TEXT2)

    for s in doc.sections:
        f = s.footer; f.is_linked_to_previous = False
        fp = f.paragraphs[0] if f.paragraphs else f.add_paragraph()
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = fp.add_run(f"Copilot Readiness — {site} — {now.strftime('%d/%m/%Y')} — CONFIDENCIAL")
        r.font.size = Pt(7); r.font.color.rgb = C_TEXT2

    doc.core_properties.title = f"Copilot Readiness - {site}"
    doc.core_properties.subject = "Prevenção de Vazamento de Dados via M365 Copilot"
    doc.core_properties.author = "sharepoint_copilot_readiness.py"
    doc.save(output)


# ============================================================
# MAIN
# ============================================================
def main():
    parser = argparse.ArgumentParser(description="Relatório DOCX de preparação para M365 Copilot.")
    parser.add_argument("excel_file", help="Ficheiro Excel de permissões.")
    parser.add_argument("-o", "--output", default=None)
    parser.add_argument("--abrir", "-a", action="store_true")
    args = parser.parse_args()
    if args.output is None: args.output = os.path.splitext(args.excel_file)[0] + "_copilot_readiness.docx"

    print("=" * 60)
    print("  🤖 Preparação M365 Copilot — Relatório de Riscos")
    print("=" * 60); print()

    print(f"[1/4] A ler: {args.excel_file}")
    data = read_excel(args.excel_file)
    t = len(data["site"]) + len(data["bibliotecas"]) + len(data["subpastas"])
    print(f"      {t} registos."); print()
    if t == 0 and not data["consolidado"]: print("ERRO: Sem dados!"); sys.exit(1)

    print("[2/4] A analisar 15 riscos Copilot...")
    engine = CopilotAuditEngine(data)
    findings, stats = engine.run()
    sc = Counter(f["sev"] for f in findings)
    print(f"      🔴 {sc.get('critico',0)} | 🟠 {sc.get('alto',0)} | 🟡 {sc.get('medio',0)} | 🔵 {sc.get('baixo',0)} | ✅ {sc.get('ok',0)}")
    print(f"      Score Copilot-Ready: {stats['copilot_score']}/100"); print()

    print(f"[3/4] A gerar DOCX...")
    generate_docx(findings, stats, data, args.excel_file, args.output)
    print(f"      {args.output} ({os.path.getsize(args.output)/1024:.0f} KB)"); print()

    print("[4/4] Concluído!"); print()
    print("=" * 60)
    print(f"  📄 {args.output}")
    print(f"  Score: {stats['copilot_score']}/100")
    if stats['copilot_score'] < 40: print("  ⛔ NÃO ATIVAR COPILOT sem corrigir os riscos!")
    elif stats['copilot_score'] < 70: print("  ⚠️  Corrigir itens críticos antes de ativar.")
    else: print("  ✅ Ambiente razoavelmente preparado.")
    print("=" * 60); print()

    print("  Riscos analisados (15):")
    for f in findings: print(f"    {f['id']}  {SEV_I.get(f['sev'],'')} {f['title']}")
    print()

    if args.abrir:
        import subprocess
        try:
            if sys.platform == "win32": os.startfile(args.output)
            elif sys.platform == "darwin": subprocess.call(["open", args.output])
            else: subprocess.call(["xdg-open", args.output])
        except: print(f"  Abra: {args.output}")

if __name__ == "__main__":
    main()
