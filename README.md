# SharePoint Governance & Copilot Readiness Toolkit

Este repositório contém um conjunto de ferramentas desenhadas para administradores de Microsoft 365 que precisam auditar permissões no SharePoint Online, garantir a conformidade com boas práticas de governança e preparar o ambiente para o uso seguro do **Microsoft 365 Copilot**.

## 🛠️ O que compõe este kit?

O toolkit é formado por três scripts complementares que trabalham em conjunto para transformar dados brutos de permissões em relatórios executivos e técnicos.

### 1. `Get-SharePointPermissions_Version3.ps1`
**Finalidade:** Extração de dados.
Este script em PowerShell utiliza o módulo PnP PowerShell para varrer um site do SharePoint (incluindo bibliotecas e subpastas) e exportar uma matriz completa de permissões para um arquivo Excel (`.xlsx`).
* **Saída:** Planilha com abas separadas por Site, Bibliotecas, Subpastas e uma visão Consolidada.

### 2. `sharepoint_audit_bestpractices_Version2.py`
**Finalidade:** Auditoria de Governança.
Script Python que lê a planilha gerada pelo primeiro passo e aplica 13 regras de boas práticas (Best Practices) do SharePoint.
* **Principais análises:** Quebras de herança excessivas, permissões diretas a usuários, contas órfãs e limites de itens por biblioteca.
* **Saída:** Relatórios detalhados em **HTML** e **DOCX**.

### 3. `sharepoint_copilot_readiness.py`
**Finalidade:** Segurança para IA (Copilot).
Focado especificamente em prevenir o vazamento de informações (oversharing) ao ativar o Microsoft 365 Copilot. Ele analisa 15 riscos críticos que podem fazer com que a IA exponha dados sensíveis a usuários não autorizados.
* **Principais análises:** Grupos amplos (Everyone/Todos), profundidade de permissões, sensibilidade de rótulos (labels) e sites candidatos ao *Restricted Content Discovery*.
* **Saída:** Relatório em **DOCX** com um **Copilot-Ready Score** (0 a 100).

## 🚀 Fluxo de Trabalho

1.  **Extração:** Execute o script PowerShell para gerar o inventário de permissões.
2.  **Auditoria:** Use o script de *Best Practices* para identificar problemas estruturais de governança.
3.  **Preparação Copilot:** Use o script de *Copilot Readiness* para validar se o ambiente está seguro para a implementação de IA Generativa.

## 📋 Pré-requisitos

* **PowerShell 7**
    * Módulo `PnP.PowerShell`
    * Módulo `ImportExcel`
* **Python 3.x**
    * Bibliotecas: `openpyxl`, `python-docx` (para os scripts Python)

## 📖 Como usar

1. **Gerar a planilha:**
   ```powershell
   .\Get-SharePointPermissions_Version3.ps1 -SiteUrl "https://suaempresa.sharepoint.com/sites/nome-do-site"
   ```

2. **Gerar auditoria de boas práticas:**
   ```bash
   python sharepoint_audit_bestpractices_Version2.py "SuaPlanilha.xlsx" -o auditoria.html
   ```

3. **Gerar relatório de prontidão Copilot:**
   ```bash
   python sharepoint_copilot_readiness.py "SuaPlanilha.xlsx" -o Copilot_Report.docx
   ```
