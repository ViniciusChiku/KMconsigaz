# 🔥 Consigaz — Auditoria Inteligente de KM por Dia

> **Desenvolvido por Chiku**  
> *Totalmente vibecodado em 2 dias para salvar o fechamento mensal e dar paz de espírito! ✨☕🚀*

---

## 💡 Sobre o Projeto

Este projeto nasceu de uma necessidade real e urgente: **acabar com a dor de cabeça e as horas infinitas conferindo cobranças manuais de KM** em planilhas de fechamento de ordens de serviço (OS) de terceirizados da Consigaz.

> [!NOTE]
> **Aviso Importante / Disclaimer:**  
> Esta ferramenta **não é** um ERP corporativo complexo nem um sistema engessado. Trata-se de uma **ferramenta prática, ágil e sob medida**, feita por quem vive a rotina do fechamento na pele, com o único objetivo de ajudar no trabalho do dia a dia e automatizar conferências manuais demoradas.  
> Foi **100% vibecodada em 2 dias** com foco total, boas xícaras de café e muita iteração rápida!

---

## ⚡ O que a Ferramenta Faz?

1. **Cruzamento Automático:**
   * Cruza a **Planilha de Fechamento** (com as abas de MOTO, CARRO, INSTALAÇÃO, CORTE/RELIGUE) com o relatório bruto extraído do **Totvs (ESOS)**.
   * Não precisa renomear colunas, alterar fórmulas nem reordenar células.

2. **Cálculo Real das Ruas (Google Maps + IBGE + OSRM):**
   * Agrupa todas as OSs atendidas no mesmo dia.
   * Monta o roteiro sequencial passando **exatamente pelas ruas visitadas**.
   * Calcula a distância rodoviária real através da malha viária oficial sem travamentos e de forma ultrarrápida.

3. **Layout "Tudo num Card só" por Dia:**
   * Cada dia com cobrança de KM possui seu próprio Card completo e independente.
   * Mostra a soma de KM cobrada no dia vs a quilometragem calculada da rota.
   * Exibe a lista completa de todas as OSs do dia com os **nomes reais e completos das ruas**.
   * Disponibiliza botões diretos de **🌐 Abrir Trecho no Google Maps** prontos para conferência visual imediata.

4. **Margem de Erro / Tolerância Customizável:**
   * Permite definir a margem de erro aceitável (ex: **40 km**).
   * Se a diferença entre o KM cobrado e o mapa estiver dentro da margem (ex: cobrou 300 km e o Maps deu 270 km), o sistema aprova automaticamente: **`✅ APROVADO (Dentro da margem de 40 km)`**.
   * Se ultrapassar a margem, ele destaca com alerta de glosa: **`🚨 EXCEDEU A MARGEM`**.

5. **Exportação Formatada em Excel:**
   * Gera um relatório consolidado com abas de resumo diário e detalhamento completo, pronto para arquivamento ou para anexar em justificativas de pagamento e glosas.

---

## 🛠️ Tecnologias Utilizadas

* **Python 3.10+**
* **Streamlit** (Interface Web moderna, responsiva e em Modo Claro)
* **OpenPyXL & MSOffCrypto** (Leitura rápida de arquivos Excel, com suporte a planilhas protegidas com senha `FECHAMENTO`)
* **OSRM & IBGE Geodata** (Cálculo de rotas rodoviárias e geolocalização dos 5.570 municípios brasileiros offline)

---

## 🚀 Como Executar

### Online (Streamlit Cloud):
Acesse diretamente pelo navegador: **[kmconsigaz.streamlit.app](https://kmconsigaz.streamlit.app)**

### Localmente:
```bash
# 1. Clone o repositório
git clone https://github.com/ViniciusChiku/KMconsigaz.git
cd KMconsigaz

# 2. Instale as dependências
pip install -r requirements.txt

# 3. Execute o aplicativo
streamlit run app_web_km.py
```

---

<p align="center">
  Feito com dedicação por <b>Chiku</b> &bull; Vibecodado em 2 dias ✨☕
</p>
