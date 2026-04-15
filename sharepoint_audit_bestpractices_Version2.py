#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Lê a planilha Excel gerada pelo Get-SharePointPermissions.ps1
e gera um relatório HTML de AUDITORIA baseado em boas práticas do SharePoint.

Analisa 13 regras de boas práticas incluindo:
  - Necessidade de dividir bibliotecas (em vez de quebrar herança)
  - Quebras de herança excessivas
  - Permissões diretas a utilizadores
  - Full Control distribuído em excesso
  - Profundidade de pastas com permissões únicas
  - Limite de 5000 permissões únicas por biblioteca
  - Grupos amplos (Everyone/Todos)
  - Permissões órfãs
  - Edição ao nível do site raiz
  - Consistência de permissões
  - Níveis personalizados
  - Pastas com público-alvo diferente do pai (candidatas a biblioteca própria)
  - Complexidade de governança

Uso:
    pip install openpyxl
    python sharepoint_audit_bestpractices.py "PermissoesSharePoint.xlsx"
    python sharepoint_audit_bestpractices.py "PermissoesSharePoint.xlsx" -o auditoria.html --abrir
"""

import argparse
import json
import os
import re
import sys
from collections import defaultdict, Counter
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("ERRO: Módulo 'openpyxl' não encontrado.")
    print("Instale com: pip install openpyxl")
    sys.exit(1)


# ============================================================
# Leitura do Excel
# ============================================================
def read_excel(filepath: str) -> dict:
    if not os.path.isfile(filepath):
        print(f"ERRO: Ficheiro não encontrado: {filepath}")
        sys.exit(1)

    wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
    data = {"resumo": {}, "site": [], "bibliotecas": [], "subpastas": [], "consolidado": []}

    sheet_map = {}
    for name in wb.sheetnames:
        lower = name.lower().strip()
        if lower == "resumo":
            sheet_map["resumo"] = name
        elif lower == "site":
            sheet_map["site"] = name
        elif lower in ("bibliotecas", "libraries"):
            sheet_map["bibliotecas"] = name
        elif lower in ("subpastas", "folders", "subfolders"):
            sheet_map["subpastas"] = name
        elif lower in ("consolidado", "consolidated", "all"):
            sheet_map["consolidado"] = name

    if "resumo" in sheet_map:
        ws = wb[sheet_map["resumo"]]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row and row[0]:
                data["resumo"][str(row[0]).strip()] = row[1] if len(row) > 1 else ""

    def read_sheet(key):
        if key not in sheet_map:
            return []
        ws = wb[sheet_map[key]]
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []
        header = [str(h).strip().lower() if h else f"col{i}" for i, h in enumerate(rows[0])]
        records = []
        for row in rows[1:]:
            if not row or all(c is None for c in row):
                continue
            record = {}
            for i, val in enumerate(row):
                if i < len(header):
                    record[header[i]] = str(val).strip() if val is not None else ""
            records.append(record)
        return records

    data["site"] = read_sheet("site")
    data["bibliotecas"] = read_sheet("bibliotecas")
    data["subpastas"] = read_sheet("subpastas")
    data["consolidado"] = read_sheet("consolidado")

    if not data["site"] and not data["bibliotecas"] and data["consolidado"]:
        for rec in data["consolidado"]:
            escopo = normalize_record(rec).get("escopo", "").lower()
            if escopo == "site":
                data["site"].append(rec)
            elif escopo in ("biblioteca", "library"):
                data["bibliotecas"].append(rec)
            elif escopo in ("pasta", "folder"):
                data["subpastas"].append(rec)

    wb.close()
    return data


def normalize_record(rec: dict) -> dict:
    field_map = {
        "escopo": "escopo", "scope": "escopo",
        "recurso": "recurso", "resource": "recurso",
        "caminhocompleto": "caminho", "caminho completo": "caminho",
        "caminho": "caminho", "path": "caminho", "fullpath": "caminho",
        "full path": "caminho",
        "permissoesunicas": "unicas", "permissões únicas": "unicas",
        "permissoesúnicas": "unicas",
        "unique": "unicas", "hasuniqueroleassignments": "unicas",
        "uniquepermissions": "unicas", "permissões unicas": "unicas",
        "tipomembro": "tipo_membro", "tipo membro": "tipo_membro",
        "membertype": "tipo_membro", "member type": "tipo_membro",
        "tipo": "tipo_membro",
        "membro": "membro", "member": "membro", "title": "membro",
        "loginmembro": "login", "login membro": "login",
        "login": "login", "loginname": "login", "memberlogin": "login",
        "niveispermissao": "niveis", "níveis permissão": "niveis",
        "niveispermissão": "niveis", "niveis permissao": "niveis",
        "níveispermissão": "niveis",
        "permissionlevels": "niveis", "permission levels": "niveis",
        "roles": "niveis", "role": "niveis", "niveis de permissao": "niveis",
        "níveis de permissão": "niveis",
    }
    normalized = {}
    for key, val in rec.items():
        clean = key.lower().strip().replace("_", "").replace("-", "")
        mapped = field_map.get(clean) or field_map.get(key.lower().strip())
        if mapped:
            normalized[mapped] = val
        else:
            normalized[clean] = val
    return normalized


def is_unique(value: str) -> bool:
    return str(value).lower().strip() in ("true", "sim", "yes", "1", "verdadeiro")


def is_user_type(tipo: str) -> bool:
    t = tipo.lower()
    return "user" in t or t == "1"


def is_full_control(niveis: str) -> bool:
    n = niveis.lower()
    return "full control" in n or "controlo total" in n or "controle total" in n


def is_edit_or_contribute(niveis: str) -> bool:
    n = niveis.lower()
    return any(x in n for x in ["edit", "contribute", "edição", "contribui", "ediç"])


def is_read_only(niveis: str) -> bool:
    n = niveis.lower()
    levels = [l.strip() for l in n.split(";") if l.strip()]
    if not levels:
        return False
    return all(
        any(r in l for r in ["read", "leitura", "view", "visuali", "apenas"])
        for l in levels
    )


# ============================================================
# Motor de Auditoria
# ============================================================
class AuditEngine:
    def __init__(self, data: dict):
        self.data = data
        self.all_records = []
        self.findings = []
        self.stats = {}

        for scope_key in ["site", "bibliotecas", "subpastas"]:
            for rec in data[scope_key]:
                n = normalize_record(rec)
                n["_scope_key"] = scope_key
                self.all_records.append(n)

    def run_all_checks(self):
        self._calc_stats()
        self._check_unique_permissions_excess()          # BP-001
        self._check_direct_user_permissions()             # BP-002
        self._check_full_control_spread()                 # BP-003
        self._check_deep_unique_permissions()             # BP-004
        self._check_library_unique_limit()                # BP-005
        self._check_everyone_permissions()                # BP-006
        self._check_orphan_like_patterns()                # BP-007
        self._check_edit_on_root_site()                   # BP-008
        self._check_permission_consistency()              # BP-009
        self._check_excessive_permission_levels()         # BP-010
        self._check_library_split_recommendation()        # BP-011 ← NOVO
        self._check_folder_audience_divergence()          # BP-012 ← NOVO
        self._check_governance_complexity_score()          # BP-013 ← NOVO

        severity_order = {"critica": 0, "alta": 1, "media": 2, "baixa": 3, "info": 4}
        self.findings.sort(key=lambda f: severity_order.get(f["severity"], 5))
        return self.findings, self.stats

    def _calc_stats(self):
        total = len(self.all_records)
        unique_count = sum(1 for r in self.all_records if is_unique(r.get("unicas", "")))
        inherited_count = total - unique_count
        members = set(r.get("membro", "") for r in self.all_records if r.get("membro"))
        direct_users = set(
            r.get("membro", "") for r in self.all_records
            if is_user_type(r.get("tipo_membro", ""))
        )
        groups = members - direct_users
        fc_count = sum(1 for r in self.all_records if is_full_control(r.get("niveis", "")))
        unique_paths = set(
            r.get("caminho", "") for r in self.all_records
            if is_unique(r.get("unicas", "")) and r.get("caminho")
        )
        lib_names = set(
            r.get("recurso", "") for r in self.all_records if r.get("_scope_key") == "bibliotecas"
        )

        # Score de saúde
        score = 100
        if total > 0:
            ur = unique_count / max(total, 1)
            dr = len(direct_users) / max(len(members), 1)
            fr = fc_count / max(total, 1)
            if ur > 0.5: score -= 25
            elif ur > 0.3: score -= 15
            elif ur > 0.1: score -= 5
            if dr > 0.5: score -= 20
            elif dr > 0.3: score -= 10
            if fr > 0.4: score -= 20
            elif fr > 0.2: score -= 10
            max_d = max((p.count("/") for p in unique_paths), default=0)
            if max_d > 5: score -= 10
            elif max_d > 3: score -= 5

        # Penalizar por bibliotecas que deveriam ser divididas
        libs_needing_split = self._count_libs_needing_split()
        if libs_needing_split > 3: score -= 15
        elif libs_needing_split > 1: score -= 8
        elif libs_needing_split > 0: score -= 3

        score = max(0, min(100, score))

        self.stats = {
            "total": total,
            "unique_count": unique_count,
            "inherited_count": inherited_count,
            "unique_ratio": round(unique_count / max(total, 1) * 100, 1),
            "total_members": len(members),
            "direct_users": len(direct_users),
            "groups": len(groups),
            "direct_user_ratio": round(len(direct_users) / max(len(members), 1) * 100, 1),
            "fc_count": fc_count,
            "fc_ratio": round(fc_count / max(total, 1) * 100, 1),
            "unique_paths": len(unique_paths),
            "lib_count": len(lib_names),
            "health_score": score,
            "direct_user_names": sorted(direct_users),
            "group_names": sorted(groups),
            "libs_needing_split": libs_needing_split,
        }

    def _get_lib_folder_data(self) -> dict:
        """Agrupa pastas com permissões únicas por biblioteca."""
        lib_folders = defaultdict(list)
        for r in self.all_records:
            if r.get("_scope_key") == "subpastas" and is_unique(r.get("unicas", "")):
                path = r.get("caminho", "")
                parts = path.split("/") if path else []
                lib_name = parts[0] if parts else r.get("recurso", "Desconhecida")
                lib_folders[lib_name].append(r)
        return lib_folders

    def _count_libs_needing_split(self) -> int:
        lib_folders = self._get_lib_folder_data()
        count = 0
        for lib, records in lib_folders.items():
            unique_first_level = set()
            for r in records:
                path = r.get("caminho", "")
                parts = path.split("/")
                if len(parts) >= 2:
                    unique_first_level.add(parts[1])
            if len(unique_first_level) >= 3:
                count += 1
        return count

    # ----------------------------------------------------------
    # BP-001: Quebras de herança
    # ----------------------------------------------------------
    def _check_unique_permissions_excess(self):
        ratio = self.stats["unique_ratio"]
        unique = self.stats["unique_count"]
        total = self.stats["total"]

        if ratio > 50: severity = "critica"
        elif ratio > 30: severity = "alta"
        elif ratio > 10: severity = "media"
        elif unique > 0: severity = "baixa"
        else: severity = "info"

        affected = []
        seen = set()
        for r in self.all_records:
            if is_unique(r.get("unicas", "")):
                path = r.get("caminho", r.get("recurso", ""))
                if path and path not in seen:
                    seen.add(path)
                    affected.append(path)

        self.findings.append({
            "id": "BP-001", "title": "Quebras de Herança de Permissões",
            "category": "Herança", "severity": severity,
            "description": (
                f"{unique} de {total} atribuições ({ratio}%) têm permissões únicas (herança quebrada). "
                + ("A Microsoft recomenda manter abaixo de 10%. Quebras excessivas tornam a gestão e auditoria praticamente impossíveis."
                   if ratio > 10 else "O nível de herança está adequado.")
            ),
            "recommendation": (
                "• Revise cada quebra de herança e avalie se é realmente necessária.\n"
                "• Em vez de quebrar herança em pastas, reorganize o conteúdo em bibliotecas separadas (ver BP-011).\n"
                "• Use grupos do SharePoint/Azure AD para gerir acesso sem quebrar herança.\n"
                "• Considere reverter para herança onde o acesso personalizado já não é necessário."
            ),
            "affected": affected[:20],
            "metric": f"{ratio}%", "metric_label": "Permissões únicas",
            "threshold": "< 10% recomendado",
        })

    # ----------------------------------------------------------
    # BP-002: Permissões diretas a utilizadores
    # ----------------------------------------------------------
    def _check_direct_user_permissions(self):
        direct_users = defaultdict(list)
        for r in self.all_records:
            if is_user_type(r.get("tipo_membro", "")):
                member = r.get("membro", "Desconhecido")
                path = r.get("caminho", r.get("recurso", ""))
                direct_users[member].append(path)

        count = len(direct_users)
        ratio = self.stats["direct_user_ratio"]

        if ratio > 50: severity = "critica"
        elif ratio > 30: severity = "alta"
        elif ratio > 10: severity = "media"
        elif count > 0: severity = "baixa"
        else: severity = "info"

        affected = []
        for user, paths in sorted(direct_users.items(), key=lambda x: -len(x[1])):
            affected.append(f"{user} → {len(paths)} locais: {', '.join(paths[:3])}{'...' if len(paths) > 3 else ''}")

        self.findings.append({
            "id": "BP-002", "title": "Permissões Diretas a Utilizadores Individuais",
            "category": "Gestão de Acesso", "severity": severity,
            "description": (
                f"{count} utilizadores têm permissões diretas ({ratio}% dos {self.stats['total_members']} membros). "
                + ("Quando alguém muda de função ou sai, é preciso revisar cada local manualmente."
                   if count > 0 else "Todas as permissões via grupos. Excelente!")
            ),
            "recommendation": (
                "• Crie grupos Azure AD ou SharePoint para cada perfil de acesso.\n"
                "• Substitua permissões individuais por pertença a grupos.\n"
                "• Facilita onboarding/offboarding: gerir grupo aplica em todos os locais.\n"
                "• Documente a finalidade de cada grupo."
            ),
            "affected": affected[:15],
            "metric": f"{count}", "metric_label": "Utilizadores diretos",
            "threshold": "0 recomendado",
        })

    # ----------------------------------------------------------
    # BP-003: Full Control
    # ----------------------------------------------------------
    def _check_full_control_spread(self):
        fc_members = defaultdict(list)
        for r in self.all_records:
            if is_full_control(r.get("niveis", "")):
                member = r.get("membro", "")
                path = r.get("caminho", r.get("recurso", ""))
                fc_members[member].append(path)

        fc_count = self.stats["fc_count"]
        fc_ratio = self.stats["fc_ratio"]

        if fc_ratio > 40: severity = "critica"
        elif fc_ratio > 25: severity = "alta"
        elif fc_ratio > 10: severity = "media"
        elif fc_count > 0: severity = "baixa"
        else: severity = "info"

        affected = [f"{m} → Full Control em {len(p)} locais"
                    for m, p in sorted(fc_members.items(), key=lambda x: -len(x[1]))]

        self.findings.append({
            "id": "BP-003", "title": "Distribuição de Full Control",
            "category": "Menor Privilégio", "severity": severity,
            "description": (
                f"{fc_count} atribuições de Full Control ({fc_ratio}%), "
                f"distribuídas por {len(fc_members)} membros/grupos. "
                + ("Full Control permite alterar permissões, eliminar conteúdo e mudar configurações."
                   if fc_ratio > 10 else "")
            ),
            "recommendation": (
                "• Princípio do Menor Privilégio: conceda apenas o mínimo necessário.\n"
                "• Substitua Full Control por 'Edit' ou 'Contribute' onde possível.\n"
                "• Reserve Full Control para administradores e proprietários do site.\n"
                "• Para quem só gere conteúdo, 'Edit' basta. Para consulta, 'Read'."
            ),
            "affected": affected[:15],
            "metric": f"{fc_ratio}%", "metric_label": "Full Control",
            "threshold": "< 10% recomendado",
        })

    # ----------------------------------------------------------
    # BP-004: Profundidade
    # ----------------------------------------------------------
    def _check_deep_unique_permissions(self):
        deep_paths = []
        for r in self.all_records:
            if is_unique(r.get("unicas", "")) and r.get("_scope_key") == "subpastas":
                path = r.get("caminho", "")
                depth = path.count("/") if path else 0
                if depth >= 3:
                    deep_paths.append((path, depth))

        deep_paths = sorted(set(deep_paths), key=lambda x: -x[1])
        max_depth = deep_paths[0][1] if deep_paths else 0

        if len(deep_paths) > 10: severity = "alta"
        elif len(deep_paths) > 3: severity = "media"
        elif deep_paths: severity = "baixa"
        else: severity = "info"

        self.findings.append({
            "id": "BP-004", "title": "Permissões Únicas em Profundidade Excessiva",
            "category": "Estrutura", "severity": severity,
            "description": (
                f"{len(deep_paths)} pastas com permissões únicas a 3+ níveis (máx: {max_depth}). "
                + ("São muito difíceis de auditar. Indicam estrutura que deve ser simplificada."
                   if deep_paths else "Sem permissões em pastas profundas.")
            ),
            "recommendation": (
                "• Evite permissões únicas além do 2.º nível de pastas.\n"
                "• Mova conteúdo sensível para bibliotecas separadas com herança (ver BP-011).\n"
                "• Utilize sites separados para projetos com requisitos de segurança distintos.\n"
                "• Use metadados e vistas em vez de pastas profundas para organizar."
            ),
            "affected": [f"[Nível {d}] {p}" for p, d in deep_paths[:15]],
            "metric": str(len(deep_paths)), "metric_label": "Pastas profundas",
            "threshold": "0 recomendado",
        })

    # ----------------------------------------------------------
    # BP-005: Limite 5000
    # ----------------------------------------------------------
    def _check_library_unique_limit(self):
        lib_unique_counts = Counter()
        for r in self.all_records:
            if is_unique(r.get("unicas", "")) and r.get("_scope_key") in ("bibliotecas", "subpastas"):
                path = r.get("caminho", "")
                lib_name = path.split("/")[0] if path else r.get("recurso", "")
                lib_unique_counts[lib_name] += 1

        warnings = []
        for lib, count in lib_unique_counts.most_common():
            if count > 5000: warnings.append((lib, count, "EXCEDIDO"))
            elif count > 3000: warnings.append((lib, count, "ATENÇÃO"))
            elif count > 1000: warnings.append((lib, count, "MONITORAR"))

        if any(w[2] == "EXCEDIDO" for w in warnings): severity = "critica"
        elif any(w[2] == "ATENÇÃO" for w in warnings): severity = "alta"
        elif warnings: severity = "media"
        else: severity = "info"

        self.findings.append({
            "id": "BP-005", "title": "Limite de Permissões Únicas por Biblioteca (5.000)",
            "category": "Performance", "severity": severity,
            "description": (
                "O SharePoint tem limite de 5.000 permissões únicas por lista/biblioteca. "
                + (f"{len(warnings)} biblioteca(s) com contagem elevada." if warnings
                   else "Nenhuma biblioteca excede limites.")
            ),
            "recommendation": (
                "• Acima de 5.000 permissões únicas, o desempenho degrada significativamente.\n"
                "• Idealmente, manter abaixo de 100 permissões únicas por biblioteca.\n"
                "• Distribua conteúdo por múltiplas bibliotecas (ver BP-011).\n"
                "• Consolide permissões usando grupos.\n"
                "• Reverta quebras de herança desnecessárias."
            ),
            "affected": [f"{lib}: {count} permissões únicas [{status}]" for lib, count, status in warnings],
            "metric": str(max(lib_unique_counts.values())) if lib_unique_counts else "0",
            "metric_label": "Máx. por biblioteca", "threshold": "< 100 ideal, < 5.000 limite",
        })

    # ----------------------------------------------------------
    # BP-006: Everyone/Todos
    # ----------------------------------------------------------
    def _check_everyone_permissions(self):
        dangerous = ["everyone", "todos", "all users", "todos os utilizadores",
                     "everyone except external", "nt authority", "all authenticated",
                     "company administrator"]
        found = []
        for r in self.all_records:
            combined = (r.get("membro", "") + " " + r.get("login", "")).lower()
            for pattern in dangerous:
                if pattern in combined:
                    path = r.get("caminho", r.get("recurso", ""))
                    found.append(f"{r.get('membro', '')} → {r.get('niveis', '')} em {path}")
                    break

        if any("full control" in f.lower() for f in found): severity = "critica"
        elif len(found) > 5: severity = "alta"
        elif found: severity = "media"
        else: severity = "info"

        self.findings.append({
            "id": "BP-006", "title": "Permissões a Grupos Amplos (Everyone / Todos)",
            "category": "Segurança", "severity": severity,
            "description": (
                f"{len(found)} atribuições a grupos amplos. "
                + ("Dão acesso a muitos utilizadores, incluindo potencialmente externos."
                   if found else "Nenhum grupo amplo. Boa prática!")
            ),
            "recommendation": (
                "• Evite 'Everyone' ou 'Everyone except external users'.\n"
                "• Substitua por grupos específicos do Azure AD.\n"
                "• Se necessário acesso amplo, use 'Read' e nunca 'Full Control'.\n"
                "• Revise regularmente quem pertence a estes grupos."
            ),
            "affected": found[:15],
            "metric": str(len(found)), "metric_label": "Atribuições amplas",
            "threshold": "0 recomendado",
        })

    # ----------------------------------------------------------
    # BP-007: Órfãs
    # ----------------------------------------------------------
    def _check_orphan_like_patterns(self):
        suspicious = []
        for r in self.all_records:
            login = r.get("login", "").lower()
            member = r.get("membro", "").lower()
            path = r.get("caminho", r.get("recurso", ""))
            if re.search(r"s-1-5-\d+-\d+", login):
                suspicious.append(f"SID não resolvido: {r.get('login', '')} em {path}")
            elif any(x in member for x in ["removed", "deleted", "removido", "eliminado"]):
                suspicious.append(f"Conta removida: {r.get('membro', '')} em {path}")
            elif login and login.startswith("i:0#.f|membership|") and "@" not in login:
                suspicious.append(f"Login suspeito: {r.get('login', '')} em {path}")

        severity = "alta" if len(suspicious) > 5 else "media" if suspicious else "info"

        self.findings.append({
            "id": "BP-007", "title": "Permissões Órfãs ou Contas Removidas",
            "category": "Limpeza", "severity": severity,
            "description": (
                f"{len(suspicious)} permissões potencialmente órfãs. "
                + ("Ocupam espaço e podem representar riscos."
                   if suspicious else "Nenhuma órfã detectada.")
            ),
            "recommendation": (
                "• Remova permissões de contas eliminadas ou SIDs não resolvidos.\n"
                "• Implemente revisão regular de acessos.\n"
                "• Use Azure AD Access Reviews para automatizar recertificação.\n"
                "• Após reestruturações, verifique permissões afetadas."
            ),
            "affected": suspicious[:15],
            "metric": str(len(suspicious)), "metric_label": "Permissões órfãs",
            "threshold": "0",
        })

    # ----------------------------------------------------------
    # BP-008: Edição no site raiz
    # ----------------------------------------------------------
    def _check_edit_on_root_site(self):
        editors = []
        for r in self.all_records:
            if r.get("_scope_key") == "site":
                niveis = r.get("niveis", "")
                if is_edit_or_contribute(niveis) or is_full_control(niveis):
                    editors.append(f"{r.get('membro', '')} ({r.get('tipo_membro', '')}) → {niveis}")

        if len(editors) > 10: severity = "alta"
        elif len(editors) > 5: severity = "media"
        elif editors: severity = "baixa"
        else: severity = "info"

        self.findings.append({
            "id": "BP-008", "title": "Permissões de Edição ao Nível do Site",
            "category": "Menor Privilégio", "severity": severity,
            "description": (
                f"{len(editors)} membros/grupos com edição ou superior ao nível do site. "
                + ("Propagam-se para TODAS as bibliotecas e pastas." if editors else "")
            ),
            "recommendation": (
                "• Limite permissões do site a 'Read' para maioria.\n"
                "• Conceda 'Edit'/'Contribute' apenas em bibliotecas/pastas específicas.\n"
                "• Reserve permissões de site para administradores.\n"
                "• Considere modelo hub/spoke: site leitura geral + bibliotecas de edição."
            ),
            "affected": editors[:15],
            "metric": str(len(editors)), "metric_label": "Editores no site",
            "threshold": "Mínimo necessário",
        })

    # ----------------------------------------------------------
    # BP-009: Inconsistência
    # ----------------------------------------------------------
    def _check_permission_consistency(self):
        member_levels = defaultdict(set)
        for r in self.all_records:
            member = r.get("membro", "")
            niveis = r.get("niveis", "")
            if member and niveis:
                member_levels[member].add(niveis)

        inconsistent = [
            f"{m}: {len(levels)} níveis → {'; '.join(sorted(levels))}"
            for m, levels in member_levels.items() if len(levels) > 2
        ]

        if len(inconsistent) > 10: severity = "media"
        elif inconsistent: severity = "baixa"
        else: severity = "info"

        self.findings.append({
            "id": "BP-009", "title": "Inconsistência de Permissões por Membro",
            "category": "Governança", "severity": severity,
            "description": (
                f"{len(inconsistent)} membros com 3+ níveis diferentes. "
                + ("Pode indicar permissões acumuladas ao longo do tempo."
                   if inconsistent else "Permissões consistentes.")
            ),
            "recommendation": (
                "• Padronize perfis de acesso com matriz de permissões.\n"
                "• Revise utilizadores com múltiplos níveis.\n"
                "• Implemente revisões trimestrais de acesso.\n"
                "• Documente: quem acede a quê e com que nível."
            ),
            "affected": inconsistent[:15],
            "metric": str(len(inconsistent)), "metric_label": "Membros inconsistentes",
            "threshold": "0 ideal",
        })

    # ----------------------------------------------------------
    # BP-010: Níveis personalizados
    # ----------------------------------------------------------
    def _check_excessive_permission_levels(self):
        all_levels = set()
        for r in self.all_records:
            for level in r.get("niveis", "").split(";"):
                level = level.strip()
                if level:
                    all_levels.add(level)

        standard = {
            "Full Control", "Design", "Edit", "Contribute", "Read",
            "Limited Access", "View Only", "Restricted Read",
            "Controlo Total", "Edição", "Contribuição", "Leitura",
            "Acesso Limitado", "Apenas Visualização",
        }
        custom = all_levels - standard

        if len(custom) > 5: severity = "media"
        elif custom: severity = "baixa"
        else: severity = "info"

        self.findings.append({
            "id": "BP-010", "title": "Níveis de Permissão Personalizados",
            "category": "Governança", "severity": severity,
            "description": (
                f"{len(all_levels)} níveis distintos, {len(custom)} personalizados. "
                + ("Muitos níveis aumentam complexidade." if custom else "Apenas padrão. Bom!")
            ),
            "recommendation": (
                "• Utilize níveis padrão sempre que possível.\n"
                "• Documente cada nível personalizado.\n"
                "• Consolide níveis similares.\n"
                "• Revise se os personalizados ainda são necessários."
            ),
            "affected": (
                [f"⚙️ Personalizado: {l}" for l in sorted(custom)] +
                [f"✅ Padrão: {l}" for l in sorted(all_levels & standard)]
            ),
            "metric": str(len(custom)), "metric_label": "Níveis personalizados",
            "threshold": "Mínimo necessário",
        })

    # ===========================================================
    # BP-011: RECOMENDAÇÃO DE DIVISÃO DE BIBLIOTECAS  ← NOVO
    # ===========================================================
    def _check_library_split_recommendation(self):
        """
        Analisa cada biblioteca e verifica se as pastas de 1.º nível com
        permissões únicas têm público-alvo (membros) suficientemente
        diferente para justificar uma biblioteca separada.

        Critérios para recomendar divisão:
        1. Biblioteca tem ≥ 3 pastas de 1.º nível com permissões únicas
        2. Público-alvo das pastas diverge significativamente (< 50% sobreposição com a biblioteca)
        3. Número total de permissões únicas na biblioteca > 100
        """
        lib_folders = self._get_lib_folder_data()

        # Obter membros da biblioteca (herança)
        lib_members = defaultdict(set)
        for r in self.all_records:
            if r.get("_scope_key") == "bibliotecas":
                recurso = r.get("recurso", "")
                lib_members[recurso].add(r.get("membro", ""))

        recommendations = []
        all_candidates = []  # Para o affected

        for lib_name, records in lib_folders.items():
            # Agrupar por pasta de 1.º nível
            first_level_folders = defaultdict(lambda: {"members": set(), "paths": set(), "count": 0})

            for r in records:
                path = r.get("caminho", "")
                parts = path.split("/")
                if len(parts) >= 2:
                    folder_key = parts[1]
                    first_level_folders[folder_key]["members"].add(r.get("membro", ""))
                    first_level_folders[folder_key]["paths"].add(path)
                    first_level_folders[folder_key]["count"] += 1

            unique_first_level = {k: v for k, v in first_level_folders.items() if v["count"] > 0}

            if len(unique_first_level) < 2:
                continue

            # Analisar sobreposição de membros entre pastas
            parent_members = lib_members.get(lib_name, set())
            total_unique_in_lib = sum(v["count"] for v in unique_first_level.values())

            divergent_folders = []
            for folder, info in unique_first_level.items():
                folder_members = info["members"]
                if parent_members:
                    overlap = len(folder_members & parent_members) / max(len(parent_members), 1)
                else:
                    overlap = 0

                # Se menos de 50% de sobreposição com a biblioteca pai,
                # ou se a pasta tem membros exclusivos
                exclusive_members = folder_members - parent_members
                if overlap < 0.5 or len(exclusive_members) > 0:
                    divergent_folders.append({
                        "folder": folder,
                        "members": sorted(folder_members),
                        "exclusive": sorted(exclusive_members),
                        "sub_paths": len(info["paths"]),
                        "overlap_pct": round(overlap * 100),
                    })

            # Decidir se recomendar divisão
            should_split = False
            reason = ""

            if len(unique_first_level) >= 3 and total_unique_in_lib > 5:
                should_split = True
                reason = (
                    f"tem {len(unique_first_level)} pastas de 1.º nível com permissões únicas "
                    f"e {total_unique_in_lib} atribuições com herança quebrada"
                )
            elif len(divergent_folders) >= 2:
                should_split = True
                reason = (
                    f"tem {len(divergent_folders)} pastas com público-alvo diferente da biblioteca"
                )
            elif total_unique_in_lib > 100:
                should_split = True
                reason = f"tem {total_unique_in_lib} permissões únicas (muito acima do ideal de 100)"

            if should_split:
                rec = {
                    "library": lib_name,
                    "reason": reason,
                    "total_unique": total_unique_in_lib,
                    "first_level_count": len(unique_first_level),
                    "divergent_count": len(divergent_folders),
                    "candidate_folders": [],
                }

                for df in divergent_folders:
                    folder_desc = (
                        f"📁 {lib_name}/{df['folder']} → "
                        f"{len(df['members'])} membros, "
                        f"{df['sub_paths']} subpastas, "
                        f"{df['overlap_pct']}% sobreposição com biblioteca"
                    )
                    if df["exclusive"]:
                        folder_desc += f"\n     Membros exclusivos: {', '.join(df['exclusive'][:5])}"
                        if len(df["exclusive"]) > 5:
                            folder_desc += f" (+{len(df['exclusive']) - 5})"
                    rec["candidate_folders"].append(folder_desc)
                    all_candidates.append(folder_desc)

                recommendations.append(rec)

        # Determinar severidade
        total_libs_to_split = len(recommendations)
        if total_libs_to_split > 3: severity = "critica"
        elif total_libs_to_split > 1: severity = "alta"
        elif total_libs_to_split > 0: severity = "media"
        else: severity = "info"

        # Construir descrição
        if recommendations:
            desc = (
                f"{total_libs_to_split} biblioteca(s) deveriam ser divididas em múltiplas bibliotecas. "
                "A Microsoft recomenda criar bibliotecas separadas em vez de quebrar herança em pastas. "
                "Cada biblioteca deve representar um perímetro de segurança onde a herança funciona naturalmente."
            )
        else:
            desc = (
                "Nenhuma biblioteca apresenta necessidade clara de divisão. "
                "A estrutura atual parece adequada ao modelo de herança."
            )

        # Affected com detalhes
        affected = []
        for rec in recommendations:
            affected.append(
                f"📚 BIBLIOTECA: {rec['library']} — {rec['reason']}"
            )
            affected.append(
                f"   → {rec['first_level_count']} pastas 1.º nível com permissões únicas, "
                f"{rec['divergent_count']} com público divergente"
            )
            affected.append("   Candidatas a tornarem-se bibliotecas próprias:")
            for cf in rec["candidate_folders"][:5]:
                affected.append(f"   {cf}")
            if len(rec["candidate_folders"]) > 5:
                affected.append(f"   ... +{len(rec['candidate_folders']) - 5} pastas")
            affected.append("")  # linha em branco

        self.findings.append({
            "id": "BP-011",
            "title": "Recomendação de Divisão de Bibliotecas",
            "category": "Arquitetura",
            "severity": severity,
            "description": desc,
            "recommendation": (
                "• Em vez de quebrar herança em pastas, crie uma biblioteca por perfil de acesso.\n"
                "• Cada pasta candidata (listada abaixo) pode tornar-se uma biblioteca independente.\n"
                "• Numa nova biblioteca, aplique permissões ao nível da biblioteca (sem quebrar herança).\n"
                "• Use metadados e vistas para organizar documentos — não pastas profundas.\n"
                "• Exemplo: 'Documentos/RH', 'Documentos/Financeiro' e 'Documentos/Jurídico'\n"
                "  com permissões únicas → criar 'Biblioteca RH', 'Biblioteca Financeiro', 'Biblioteca Jurídico'.\n"
                "• Após migração, verifique que a herança funciona sem quebras.\n"
                "• Documente a nova estrutura e comunique às equipas."
            ),
            "affected": affected,
            "metric": str(total_libs_to_split),
            "metric_label": "Bibliotecas a dividir",
            "threshold": "0",
        })

    # ===========================================================
    # BP-012: DIVERGÊNCIA DE PÚBLICO-ALVO EM PASTAS  ← NOVO
    # ===========================================================
    def _check_folder_audience_divergence(self):
        """
        Para cada pasta com permissões únicas, compara os membros dessa pasta
        com os membros da biblioteca pai. Identifica pastas onde o público é
        completamente diferente — forte indicador de que deveria ser noutra
        biblioteca ou site.
        """
        lib_members = defaultdict(set)
        for r in self.all_records:
            if r.get("_scope_key") == "bibliotecas":
                lib_members[r.get("recurso", "")].add(r.get("membro", ""))

        # Agrupar pastas por caminho
        folder_members = defaultdict(set)
        folder_lib_map = {}
        for r in self.all_records:
            if r.get("_scope_key") == "subpastas" and is_unique(r.get("unicas", "")):
                path = r.get("caminho", "")
                folder_members[path].add(r.get("membro", ""))
                parts = path.split("/")
                if parts:
                    folder_lib_map[path] = parts[0]

        high_divergence = []
        for path, members in folder_members.items():
            lib = folder_lib_map.get(path, "")
            parent = lib_members.get(lib, set())
            if not parent or not members:
                continue

            # Calcular Jaccard distance
            intersection = len(members & parent)
            union = len(members | parent)
            similarity = intersection / max(union, 1)

            if similarity < 0.3:  # Menos de 30% sobreposição
                exclusive = members - parent
                high_divergence.append({
                    "path": path,
                    "library": lib,
                    "similarity": round(similarity * 100),
                    "folder_members": sorted(members),
                    "exclusive_members": sorted(exclusive),
                })

        high_divergence.sort(key=lambda x: x["similarity"])

        if len(high_divergence) > 5: severity = "alta"
        elif len(high_divergence) > 2: severity = "media"
        elif high_divergence: severity = "baixa"
        else: severity = "info"

        affected = []
        for hd in high_divergence[:15]:
            affected.append(
                f"📁 {hd['path']} (em '{hd['library']}') → "
                f"apenas {hd['similarity']}% sobreposição com biblioteca"
            )
            if hd["exclusive_members"]:
                exc = ", ".join(hd["exclusive_members"][:4])
                if len(hd["exclusive_members"]) > 4:
                    exc += f" (+{len(hd['exclusive_members']) - 4})"
                affected.append(f"   Membros exclusivos da pasta: {exc}")

        self.findings.append({
            "id": "BP-012",
            "title": "Pastas com Público-Alvo Divergente da Biblioteca",
            "category": "Arquitetura",
            "severity": severity,
            "description": (
                f"{len(high_divergence)} pasta(s) têm público-alvo com menos de 30% de sobreposição "
                "com a biblioteca pai. "
                + ("Isto indica que o conteúdo está mal posicionado — estas pastas servem um público "
                   "completamente diferente e deveriam estar numa biblioteca ou site próprio."
                   if high_divergence else "Todos os públicos estão alinhados com as bibliotecas.")
            ),
            "recommendation": (
                "• Cada pasta com público divergente é candidata a ter a sua própria biblioteca.\n"
                "• Crie uma biblioteca dedicada e mova o conteúdo.\n"
                "• Na nova biblioteca, atribua permissões por herança (sem quebrar).\n"
                "• Se o público for de outra equipa/departamento, considere um site separado.\n"
                "• Utilize o SharePoint Hub para ligar sites de diferentes equipas."
            ),
            "affected": affected,
            "metric": str(len(high_divergence)),
            "metric_label": "Pastas divergentes",
            "threshold": "0",
        })

    # ===========================================================
    # BP-013: COMPLEXIDADE DE GOVERNANÇA  ← NOVO
    # ===========================================================
    def _check_governance_complexity_score(self):
        """
        Calcula um índice de complexidade de governança baseado em múltiplos fatores:
        - Número de permissões únicas
        - Utilizadores diretos
        - Profundidade de herança quebrada
        - Número de níveis de permissão diferentes
        - Número de bibliotecas com quebras
        Fornece uma avaliação holística e plano de ação priorizado.
        """
        unique = self.stats["unique_count"]
        total = self.stats["total"]
        direct = self.stats["direct_users"]
        fc = self.stats["fc_count"]
        libs_split = self.stats.get("libs_needing_split", 0)

        # Calcular profundidade máxima
        max_depth = 0
        for r in self.all_records:
            if is_unique(r.get("unicas", "")) and r.get("_scope_key") == "subpastas":
                d = r.get("caminho", "").count("/")
                if d > max_depth:
                    max_depth = d

        # Número de níveis distintos
        all_levels = set()
        for r in self.all_records:
            for l in r.get("niveis", "").split(";"):
                if l.strip():
                    all_levels.add(l.strip())

        # Calcular complexidade (0-100, onde 0 = simples, 100 = caótico)
        complexity = 0
        factors = []

        # Fator 1: Ratio de permissões únicas (0-25 pontos)
        ur = unique / max(total, 1)
        f1 = min(25, int(ur * 50))
        complexity += f1
        if f1 > 10: factors.append(f"Herança quebrada: {self.stats['unique_ratio']}% (+{f1} pts)")

        # Fator 2: Utilizadores diretos (0-20 pontos)
        dr = direct / max(self.stats["total_members"], 1)
        f2 = min(20, int(dr * 40))
        complexity += f2
        if f2 > 5: factors.append(f"Utilizadores diretos: {direct} (+{f2} pts)")

        # Fator 3: Full Control (0-15 pontos)
        fr = fc / max(total, 1)
        f3 = min(15, int(fr * 30))
        complexity += f3
        if f3 > 5: factors.append(f"Full Control: {self.stats['fc_ratio']}% (+{f3} pts)")

        # Fator 4: Profundidade (0-15 pontos)
        f4 = min(15, max_depth * 3)
        complexity += f4
        if f4 > 3: factors.append(f"Profundidade máxima: {max_depth} níveis (+{f4} pts)")

        # Fator 5: Bibliotecas para dividir (0-15 pontos)
        f5 = min(15, libs_split * 5)
        complexity += f5
        if f5 > 0: factors.append(f"Bibliotecas a dividir: {libs_split} (+{f5} pts)")

        # Fator 6: Níveis de permissão (0-10 pontos)
        f6 = min(10, max(0, len(all_levels) - 5) * 2)
        complexity += f6
        if f6 > 2: factors.append(f"Níveis de permissão: {len(all_levels)} (+{f6} pts)")

        complexity = min(100, complexity)

        if complexity > 70: severity = "critica"
        elif complexity > 50: severity = "alta"
        elif complexity > 30: severity = "media"
        elif complexity > 10: severity = "baixa"
        else: severity = "info"

        # Plano de ação priorizado
        actions = []
        if libs_split > 0:
            actions.append(f"1.º DIVIDIR BIBLIOTECAS: {libs_split} biblioteca(s) devem ser reorganizadas em múltiplas (BP-011)")
        if ur > 0.1:
            actions.append(f"2.º REDUZIR QUEBRAS: Reverter herança desnecessária ({unique} atribuições únicas)")
        if direct > 0:
            actions.append(f"3.º MIGRAR PARA GRUPOS: Substituir {direct} permissões diretas por grupos Azure AD/SP")
        if fr > 0.1:
            actions.append(f"4.º REDUZIR FULL CONTROL: Rebaixar {fc} atribuições para Edit/Contribute/Read")
        if max_depth > 3:
            actions.append(f"5.º SIMPLIFICAR ESTRUTURA: Reduzir profundidade de {max_depth} para máx. 2 níveis")
        if not actions:
            actions.append("✅ A estrutura de permissões está bem governada!")

        self.findings.append({
            "id": "BP-013",
            "title": "Índice de Complexidade de Governança",
            "category": "Visão Geral",
            "severity": severity,
            "description": (
                f"Índice de complexidade: {complexity}/100 "
                f"({'Simples' if complexity <= 20 else 'Moderada' if complexity <= 40 else 'Elevada' if complexity <= 60 else 'Muito elevada' if complexity <= 80 else 'Caótica'}). "
                "Este índice combina todos os fatores de risco para dar uma visão holística da "
                "complexidade de gestão das permissões deste site."
            ),
            "recommendation": (
                "PLANO DE AÇÃO PRIORIZADO:\n\n" +
                "\n".join(actions) +
                "\n\n"
                "NOTA: A ordem é estratégica — reorganizar bibliotecas (BP-011) primeiro "
                "porque frequentemente resolve múltiplos problemas de uma vez."
            ),
            "affected": [f"Fator: {f}" for f in factors] if factors else ["Todos os fatores dentro dos limites."],
            "metric": str(complexity),
            "metric_label": "Complexidade (0-100)",
            "threshold": "< 20 ideal",
        })


# ============================================================
# Gerar HTML
# ============================================================
def generate_audit_html(findings: list, stats: dict, data: dict, excel_filename: str) -> str:
    resumo = data["resumo"]
    site_name = resumo.get("Site", resumo.get("site", "SharePoint Site"))
    site_url = resumo.get("URL", resumo.get("url", ""))
    now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    score = stats["health_score"]
    if score >= 80: score_color, score_label, score_icon = "#10b981", "Saudável", "&#x2705;"
    elif score >= 60: score_color, score_label, score_icon = "#f59e0b", "Atenção Necessária", "&#x26A0;&#xFE0F;"
    elif score >= 40: score_color, score_label, score_icon = "#f97316", "Risco Moderado", "&#x1F7E0;"
    else: score_color, score_label, score_icon = "#ef4444", "Risco Elevado", "&#x1F534;"

    sev_counts = Counter(f["severity"] for f in findings)
    criticas = sev_counts.get("critica", 0)
    altas = sev_counts.get("alta", 0)
    medias = sev_counts.get("media", 0)
    baixas = sev_counts.get("baixa", 0)
    infos = sev_counts.get("info", 0)

    findings_json = json.dumps(findings, ensure_ascii=False)
    stats_json = json.dumps(stats, ensure_ascii=False)

    html = f"""<!DOCTYPE html>
<html lang="pt">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Auditoria SharePoint - {site_name}</title>
    <style>
        :root {{
            --bg-primary:#0f172a; --bg-secondary:#1e293b; --bg-card:#1e293b;
            --bg-card-hover:#334155; --text-primary:#f1f5f9; --text-secondary:#94a3b8;
            --text-muted:#64748b; --accent-blue:#3b82f6; --accent-green:#10b981;
            --accent-orange:#f59e0b; --accent-red:#ef4444; --accent-purple:#8b5cf6;
            --accent-cyan:#06b6d4; --border-color:#334155;
            --sev-critica:#ef4444; --sev-alta:#f97316; --sev-media:#f59e0b;
            --sev-baixa:#3b82f6; --sev-info:#10b981;
        }}
        *{{margin:0;padding:0;box-sizing:border-box}}
        body{{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,'Helvetica Neue',Arial,sans-serif;background:var(--bg-primary);color:var(--text-primary);line-height:1.6}}

        .header{{background:linear-gradient(135deg,#1e3a5f 0%,#0f172a 100%);border-bottom:1px solid var(--border-color);padding:24px 32px}}
        .header h1{{font-size:22px;font-weight:600;display:flex;align-items:center;gap:10px}}
        .header .subtitle{{color:var(--text-secondary);font-size:13px;margin-top:4px}}
        .header .subtitle a{{color:var(--accent-blue);text-decoration:none}}
        .header-meta{{display:flex;gap:16px;margin-top:8px;font-size:12px;color:var(--text-muted);flex-wrap:wrap}}

        .score-hero{{display:flex;align-items:center;justify-content:center;gap:40px;padding:32px;flex-wrap:wrap}}
        .score-circle{{width:180px;height:180px;border-radius:50%;background:conic-gradient({score_color} {score*3.6}deg,#334155 0deg);display:flex;align-items:center;justify-content:center}}
        .score-circle-inner{{width:150px;height:150px;border-radius:50%;background:var(--bg-primary);display:flex;flex-direction:column;align-items:center;justify-content:center}}
        .score-number{{font-size:48px;font-weight:700;color:{score_color};line-height:1}}
        .score-of{{font-size:12px;color:var(--text-muted)}}
        .score-info{{max-width:420px}}
        .score-label{{font-size:24px;font-weight:600;color:{score_color};margin-bottom:8px}}
        .score-desc{{color:var(--text-secondary);font-size:14px;line-height:1.6}}
        .severity-summary{{display:flex;gap:10px;margin-top:16px;flex-wrap:wrap}}
        .sev-pill{{padding:6px 14px;border-radius:20px;font-size:12px;font-weight:600;display:flex;align-items:center;gap:6px}}
        .sev-pill.critica{{background:rgba(239,68,68,.15);color:#f87171;border:1px solid rgba(239,68,68,.3)}}
        .sev-pill.alta{{background:rgba(249,115,22,.15);color:#fb923c;border:1px solid rgba(249,115,22,.3)}}
        .sev-pill.media{{background:rgba(245,158,11,.15);color:#fbbf24;border:1px solid rgba(245,158,11,.3)}}
        .sev-pill.baixa{{background:rgba(59,130,246,.15);color:#60a5fa;border:1px solid rgba(59,130,246,.3)}}
        .sev-pill.info{{background:rgba(16,185,129,.15);color:#34d399;border:1px solid rgba(16,185,129,.3)}}

        .metrics{{display:grid;grid-template-columns:repeat(auto-fit,minmax(155px,1fr));gap:14px;padding:0 32px 24px}}
        .metric-card{{background:var(--bg-card);border:1px solid var(--border-color);border-radius:10px;padding:16px;text-align:center;transition:all .2s}}
        .metric-card:hover{{border-color:var(--accent-blue);transform:translateY(-2px)}}
        .metric-card .val{{font-size:26px;font-weight:700;background:linear-gradient(135deg,var(--accent-blue),var(--accent-cyan));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}}
        .metric-card.warn .val{{background:linear-gradient(135deg,var(--accent-red),var(--accent-orange));-webkit-background-clip:text;-webkit-text-fill-color:transparent;background-clip:text}}
        .metric-card .lbl{{color:var(--text-secondary);font-size:11px;text-transform:uppercase;letter-spacing:.5px;margin-top:2px}}
        .metric-card .thr{{color:var(--text-muted);font-size:10px;margin-top:4px}}

        .findings-container{{padding:0 32px 32px}}
        .findings-header{{display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;flex-wrap:wrap;gap:12px}}
        .findings-header h2{{font-size:18px;font-weight:600}}
        .filter-btns{{display:flex;gap:6px;flex-wrap:wrap}}
        .filter-btn{{padding:6px 12px;border-radius:6px;border:1px solid var(--border-color);background:var(--bg-secondary);color:var(--text-secondary);font-size:12px;cursor:pointer;transition:all .2s}}
        .filter-btn:hover{{border-color:var(--accent-blue)}}
        .filter-btn.active{{background:rgba(59,130,246,.15);border-color:var(--accent-blue);color:var(--accent-blue)}}

        .finding-card{{background:var(--bg-card);border:1px solid var(--border-color);border-radius:12px;margin-bottom:16px;overflow:hidden;transition:all .2s;border-left:4px solid var(--border-color)}}
        .finding-card:hover{{border-color:rgba(59,130,246,.4)}}
        .finding-card.sev-critica{{border-left-color:var(--sev-critica)}}
        .finding-card.sev-alta{{border-left-color:var(--sev-alta)}}
        .finding-card.sev-media{{border-left-color:var(--sev-media)}}
        .finding-card.sev-baixa{{border-left-color:var(--sev-baixa)}}
        .finding-card.sev-info{{border-left-color:var(--sev-info)}}

        .finding-header{{padding:16px 20px;cursor:pointer;display:flex;align-items:center;gap:12px;user-select:none}}
        .finding-header:hover{{background:var(--bg-card-hover)}}
        .finding-toggle{{font-size:10px;color:var(--text-muted);transition:transform .2s;width:20px;flex-shrink:0;text-align:center}}
        .finding-header.expanded .finding-toggle{{transform:rotate(90deg)}}
        .finding-sev{{font-size:10px;padding:3px 10px;border-radius:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;flex-shrink:0}}
        .finding-sev.critica{{background:rgba(239,68,68,.15);color:#f87171}}
        .finding-sev.alta{{background:rgba(249,115,22,.15);color:#fb923c}}
        .finding-sev.media{{background:rgba(245,158,11,.15);color:#fbbf24}}
        .finding-sev.baixa{{background:rgba(59,130,246,.15);color:#60a5fa}}
        .finding-sev.info{{background:rgba(16,185,129,.15);color:#34d399}}
        .finding-id{{font-size:11px;color:var(--text-muted);font-family:monospace;flex-shrink:0}}
        .finding-title{{font-size:14px;font-weight:600;flex:1}}
        .finding-cat{{font-size:10px;padding:2px 8px;border-radius:4px;background:rgba(139,92,246,.15);color:#a78bfa;border:1px solid rgba(139,92,246,.2);flex-shrink:0}}
        .finding-metric{{text-align:right;flex-shrink:0;min-width:80px}}
        .finding-metric .val{{font-size:18px;font-weight:700}}
        .finding-metric .lbl{{font-size:10px;color:var(--text-muted)}}

        .finding-body{{display:none;padding:0 20px 20px;border-top:1px solid var(--border-color)}}
        .finding-body.expanded{{display:block}}
        .finding-section{{margin-top:16px}}
        .finding-section h4{{font-size:12px;text-transform:uppercase;letter-spacing:.5px;color:var(--text-muted);margin-bottom:8px;display:flex;align-items:center;gap:6px}}
        .finding-desc{{color:var(--text-secondary);font-size:13px;line-height:1.7}}
        .finding-rec{{background:rgba(16,185,129,.05);border:1px solid rgba(16,185,129,.15);border-radius:8px;padding:14px 16px;font-size:13px;color:var(--text-secondary);line-height:1.8;white-space:pre-line}}
        .finding-affected{{background:rgba(239,68,68,.03);border:1px solid rgba(239,68,68,.1);border-radius:8px;padding:12px 16px;max-height:300px;overflow-y:auto;font-size:12px;font-family:'Cascadia Code','Fira Code',Consolas,monospace;color:var(--text-secondary);line-height:1.8}}
        .finding-affected .item{{padding:3px 0;border-bottom:1px solid rgba(51,65,85,.3)}}
        .finding-affected .item:last-child{{border-bottom:none}}

        .members-section{{padding:0 32px 32px}}
        .members-section h2{{font-size:18px;font-weight:600;margin-bottom:16px}}
        .members-grid{{display:grid;grid-template-columns:repeat(auto-fill,minmax(280px,1fr));gap:12px}}
        .member-card{{background:var(--bg-card);border:1px solid var(--border-color);border-radius:10px;padding:14px 16px;display:flex;align-items:center;gap:12px;transition:all .2s}}
        .member-card:hover{{border-color:var(--accent-blue)}}
        .member-icon{{font-size:24px}}
        .member-info{{flex:1;min-width:0}}
        .member-name{{font-size:13px;font-weight:500;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}}
        .member-type{{font-size:10px;padding:1px 6px;border-radius:4px;display:inline-block;margin-top:2px}}
        .member-type.direct{{background:rgba(239,68,68,.15);color:#f87171}}
        .member-type.group{{background:rgba(16,185,129,.15);color:#34d399}}

        .footer{{text-align:center;padding:24px;color:var(--text-muted);font-size:11px;border-top:1px solid var(--border-color);margin-top:16px}}

        @media(max-width:768px){{
            .header,.score-hero,.metrics,.findings-container,.members-section{{padding-left:16px;padding-right:16px}}
            .score-hero{{flex-direction:column;gap:20px}}
        }}
    </style>
</head>
<body>
<div class="header">
    <h1><span>&#x1F50D;</span> Auditoria de Permissões &mdash; Boas Práticas SharePoint</h1>
    <div class="subtitle">{"<a href='"+site_url+"' target='_blank'>"+site_url+"</a>" if site_url else site_name}</div>
    <div class="header-meta">
        <span>&#x1F4C5; {now}</span>
        <span>&#x1F4C2; Fonte: {os.path.basename(excel_filename)}</span>
        <span>&#x1F4CA; {stats["total"]} atribuições analisadas</span>
    </div>
</div>

<div class="score-hero">
    <div class="score-circle"><div class="score-circle-inner">
        <div class="score-number">{score}</div><div class="score-of">de 100</div>
    </div></div>
    <div class="score-info">
        <div class="score-label">{score_icon} {score_label}</div>
        <div class="score-desc">
            Pontuação de aderência às boas práticas de permissões do SharePoint,
            considerando herança, grupos, menor privilégio, estrutura e necessidade de reorganização de bibliotecas.
        </div>
        <div class="severity-summary">
            <span class="sev-pill critica">&#x1F534; {criticas} Crítica{"s" if criticas!=1 else ""}</span>
            <span class="sev-pill alta">&#x1F7E0; {altas} Alta{"s" if altas!=1 else ""}</span>
            <span class="sev-pill media">&#x1F7E1; {medias} Média{"s" if medias!=1 else ""}</span>
            <span class="sev-pill baixa">&#x1F535; {baixas} Baixa{"s" if baixas!=1 else ""}</span>
            <span class="sev-pill info">&#x2705; {infos} OK</span>
        </div>
    </div>
</div>

<div class="metrics">
    <div class="metric-card {"warn" if stats["unique_ratio"]>30 else ""}">
        <div class="val">{stats["unique_ratio"]}%</div><div class="lbl">Permissões Únicas</div><div class="thr">&lt; 10% recomendado</div>
    </div>
    <div class="metric-card {"warn" if stats["direct_user_ratio"]>30 else ""}">
        <div class="val">{stats["direct_users"]}</div><div class="lbl">Utilizadores Diretos</div><div class="thr">0 recomendado</div>
    </div>
    <div class="metric-card {"warn" if stats["fc_ratio"]>25 else ""}">
        <div class="val">{stats["fc_ratio"]}%</div><div class="lbl">Full Control</div><div class="thr">&lt; 10% recomendado</div>
    </div>
    <div class="metric-card {"warn" if stats["libs_needing_split"]>0 else ""}">
        <div class="val">{stats["libs_needing_split"]}</div><div class="lbl">Bibliotecas a Dividir</div><div class="thr">0 ideal</div>
    </div>
    <div class="metric-card">
        <div class="val">{stats["total_members"]}</div><div class="lbl">Membros Totais</div><div class="thr">{stats["groups"]} grupos, {stats["direct_users"]} diretos</div>
    </div>
    <div class="metric-card">
        <div class="val">{stats["unique_paths"]}</div><div class="lbl">Locais c/ Perm. Única</div><div class="thr">Mínimo necessário</div>
    </div>
    <div class="metric-card">
        <div class="val">{stats["lib_count"]}</div><div class="lbl">Bibliotecas</div><div class="thr">&nbsp;</div>
    </div>
</div>

<div class="findings-container">
    <div class="findings-header">
        <h2>&#x1F4CB; Resultados da Auditoria ({len(findings)} verificações)</h2>
        <div class="filter-btns">
            <button class="filter-btn active" onclick="filterFindings('all',this)">Todas</button>
            <button class="filter-btn" onclick="filterFindings('critica',this)">&#x1F534; Críticas</button>
            <button class="filter-btn" onclick="filterFindings('alta',this)">&#x1F7E0; Altas</button>
            <button class="filter-btn" onclick="filterFindings('media',this)">&#x1F7E1; Médias</button>
            <button class="filter-btn" onclick="filterFindings('baixa',this)">&#x1F535; Baixas</button>
            <button class="filter-btn" onclick="filterFindings('info',this)">&#x2705; OK</button>
            <button class="filter-btn" onclick="filterFindings('arquitetura',this)">&#x1F3D7; Arquitetura</button>
        </div>
    </div>
    <div id="findingsContainer"></div>
</div>

<div class="members-section">
    <h2>&#x1F465; Visão Geral de Membros</h2>
    <div class="members-grid" id="membersGrid"></div>
</div>

<div class="footer">
    Relatório de auditoria gerado por <strong>sharepoint_audit_bestpractices.py</strong>
    a partir de <strong>{os.path.basename(excel_filename)}</strong> em {now}.<br>
    Baseado nas boas práticas da Microsoft para SharePoint Online (2025/2026).
</div>

<script>
const FINDINGS={findings_json};
const STATS={stats_json};

function renderFindings(){{
    const c=document.getElementById('findingsContainer');c.innerHTML='';
    FINDINGS.forEach((f,i)=>{{
        const card=document.createElement('div');
        card.className='finding-card sev-'+f.severity;
        card.dataset.severity=f.severity;
        card.dataset.category=(f.category||'').toLowerCase();
        const sl={{critica:'CRÍTICA',alta:'ALTA',media:'MÉDIA',baixa:'BAIXA',info:'OK'}}[f.severity]||'';
        const si={{critica:'\\uD83D\\uDD34',alta:'\\uD83D\\uDFE0',media:'\\uD83D\\uDFE1',baixa:'\\uD83D\\uDD35',info:'\\u2705'}}[f.severity]||'';
        card.innerHTML=`
            <div class="finding-header" onclick="toggleFinding(${{i}})">
                <span class="finding-toggle" id="toggle-${{i}}">\\u25B6</span>
                <span class="finding-sev ${{f.severity}}">${{si}} ${{sl}}</span>
                <span class="finding-id">${{f.id}}</span>
                <span class="finding-title">${{f.title}}</span>
                <span class="finding-cat">${{f.category}}</span>
                <div class="finding-metric"><div class="val">${{f.metric}}</div><div class="lbl">${{f.metric_label}}</div></div>
            </div>
            <div class="finding-body" id="body-${{i}}">
                <div class="finding-section"><h4>\\uD83D\\uDCDD Diagnóstico</h4><div class="finding-desc">${{esc(f.description)}}</div></div>
                <div class="finding-section"><h4>\\u2705 Recomendação</h4><div class="finding-rec">${{esc(f.recommendation)}}</div></div>
                ${{f.affected&&f.affected.length?`<div class="finding-section"><h4>\\uD83D\\uDCCD Itens Afetados (${{f.affected.length}}${{f.affected.length>=15?'+':''}})</h4><div class="finding-affected">${{f.affected.map(a=>'<div class="item">'+esc(a)+'</div>').join('')}}</div></div>`:''}}
                <div class="finding-section" style="margin-top:12px"><span style="font-size:11px;color:var(--text-muted)">Threshold: ${{f.threshold}}</span></div>
            </div>`;
        c.appendChild(card);
    }});
}}

function toggleFinding(i){{
    const h=document.getElementById('body-'+i).previousElementSibling;
    const b=document.getElementById('body-'+i);
    const o=b.classList.contains('expanded');
    if(o){{b.classList.remove('expanded');h.classList.remove('expanded')}}
    else{{b.classList.add('expanded');h.classList.add('expanded')}}
}}

function filterFindings(sev,btn){{
    document.querySelectorAll('.filter-btn').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    document.querySelectorAll('.finding-card').forEach(card=>{{
        if(sev==='all'){{card.style.display=''}}
        else if(sev==='arquitetura'){{card.style.display=card.dataset.category==='arquitetura'?'':'none'}}
        else{{card.style.display=card.dataset.severity===sev?'':'none'}}
    }});
}}

function esc(s){{return(s||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')}}

function renderMembers(){{
    const g=document.getElementById('membersGrid');g.innerHTML='';
    const du=STATS.direct_user_names||[];const gr=STATS.group_names||[];
    du.forEach(n=>{{g.innerHTML+=`<div class="member-card"><div class="member-icon">\\uD83D\\uDC64</div><div class="member-info"><div class="member-name">${{esc(n)}}</div><span class="member-type direct">\\u26A0 Permissão direta</span></div></div>`}});
    gr.forEach(n=>{{g.innerHTML+=`<div class="member-card"><div class="member-icon">\\uD83D\\uDC65</div><div class="member-info"><div class="member-name">${{esc(n)}}</div><span class="member-type group">\\u2705 Grupo</span></div></div>`}});
    if(!du.length&&!gr.length)g.innerHTML='<p style="color:var(--text-muted)">Nenhum membro encontrado.</p>';
}}

function autoExpand(){{
    FINDINGS.forEach((f,i)=>{{
        if(f.severity==='critica'||f.severity==='alta'){{
            const b=document.getElementById('body-'+i);const h=b.previousElementSibling;
            b.classList.add('expanded');h.classList.add('expanded');
        }}
    }});
}}

renderFindings();renderMembers();autoExpand();
</script>
</body>
</html>"""
    return html


# ============================================================
# MAIN
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description="Gera relatório HTML de auditoria de permissões SharePoint baseado em boas práticas."
    )
    parser.add_argument("excel_file", help="Ficheiro Excel (.xlsx) de permissões.")
    parser.add_argument("-o", "--output", default=None, help="Ficheiro HTML de saída.")
    parser.add_argument("--abrir", "-a", action="store_true", help="Abrir no browser.")

    args = parser.parse_args()
    if args.output is None:
        args.output = os.path.splitext(args.excel_file)[0] + "_auditoria.html"

    print("=" * 58)
    print("  🔍 Auditoria SharePoint — Boas Práticas (v2)")
    print("=" * 58)
    print()

    print(f"[1/4] A ler: {args.excel_file}")
    data = read_excel(args.excel_file)
    total = len(data["site"]) + len(data["bibliotecas"]) + len(data["subpastas"])
    print(f"      {total} registos carregados.")
    print()

    if total == 0 and not data["consolidado"]:
        print("ERRO: Nenhum dado encontrado!")
        sys.exit(1)

    print("[2/4] A executar 13 verificações de boas práticas...")
    engine = AuditEngine(data)
    findings, stats = engine.run_all_checks()

    sev = Counter(f["severity"] for f in findings)
    print(f"      🔴 Críticas: {sev.get('critica',0)}")
    print(f"      🟠 Altas:    {sev.get('alta',0)}")
    print(f"      🟡 Médias:   {sev.get('media',0)}")
    print(f"      🔵 Baixas:   {sev.get('baixa',0)}")
    print(f"      ✅ OK:       {sev.get('info',0)}")
    print(f"      Score: {stats['health_score']}/100")
    print(f"      Bibliotecas a dividir: {stats['libs_needing_split']}")
    print()

    print(f"[3/4] A gerar HTML...")
    html = generate_audit_html(findings, stats, data, args.excel_file)
    with open(args.output, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"      {args.output} ({os.path.getsize(args.output)/1024:.1f} KB)")
    print()

    print("[4/4] Concluído!")
    print()
    print("=" * 58)
    print(f"  Score: {stats['health_score']}/100")
    print("=" * 58)
    print()
    print("  Verificações (13):")
    print("    BP-001  Quebras de herança de permissões")
    print("    BP-002  Permissões diretas a utilizadores")
    print("    BP-003  Distribuição de Full Control")
    print("    BP-004  Permissões únicas em profundidade")
    print("    BP-005  Limite 5.000 por biblioteca")
    print("    BP-006  Grupos amplos (Everyone/Todos)")
    print("    BP-007  Permissões órfãs / contas removidas")
    print("    BP-008  Edição ao nível do site raiz")
    print("    BP-009  Inconsistência de permiss��es")
    print("    BP-010  Níveis de permissão personalizados")
    print("    BP-011  ⭐ Recomendação de divisão de bibliotecas")
    print("    BP-012  ⭐ Divergência de público-alvo em pastas")
    print("    BP-013  ⭐ Índice de complexidade de governança")
    print()

    if args.abrir:
        import webbrowser
        webbrowser.open(args.output)
        print("  🌐 A abrir no browser...")


if __name__ == "__main__":
    main()