# B2B Lead Generation & Enrichment Pipeline

Pipeline modular em Python para extração, consolidação, pontuação e enriquecimento de leads B2B a partir de bases públicas (dados abertos de CNPJ da Receita Federal e OpenStreetMap)[cite: 11, 12, 13, 14].

## Etapas do Pipeline

1. **Extração via base aberta de CNPJ (`leads_cnpj_receita.py`)**[cite: 11]
   - Lê arquivos `.zip` da Receita Federal (`Estabelecimentos`, `Empresas` e `Municipios`) em *chunks* de 500 mil registros para controle de memória[cite: 11].
   - Filtra estabelecimentos ativos por UF, município e CNAE (principal e/ou secundário) e gera CSV inicial com pontuação base[cite: 11].

2. **Consolidação e *lead scoring* (`consolida_leads.py`)**[cite: 13]
   - Agrupa estabelecimentos por CNPJ básico (8 dígitos) para consolidar empresas e evitar duplicidade de filiais[cite: 13].
   - Identifica redes nacionais, distingue matrizes/filiais e sinaliza postos de coleta e registros que exigem revisão manual[cite: 13].
   - Classifica e-mails (`próprio`, `compartilhado`, `gratuito`, `administrativo`, `sem e-mail`), calcula *score* (0–100) e prioridade (`A`, `B`, `C`) e exporta planilha `.xlsx` formatada + `.csv` de empresas[cite: 13].

3. **Enriquecimento web (`enriquece_leads.py`)**[cite: 14]
   - Identifica sites a partir do domínio do e-mail ou, opcionalmente, via Google Places API (`--places-key`) com verificação de similaridade de nome e cidade[cite: 14].
   - Varre a *home* e páginas institucionais/de contato respeitando `robots.txt`, *rate limit* por *host* e execução concorrente (`ThreadPoolExecutor`)[cite: 14].
   - Extrai WhatsApp, e-mails (priorizando áreas comerciais/compras/diretoria e filtrando endereços administrativos/ruído), telefones, redes sociais, menções a marcas e possível responsável técnico[cite: 14].
   - Salva progresso incremental em cache `.jsonl` para retomada em caso de interrupção[cite: 14].

4. **Coleta alternativa via OpenStreetMap (`leads_laboratorios.py`)**[cite: 12]
   - Consulta a Overpass API por estado ou município e varre páginas de contato dos sites encontrados para consolidar canais públicos[cite: 12].

## Como Executar

Instale as dependências necessárias:
```powershell
pip install pandas openpyxl requests beautifulsoup4
```[cite: 12, 13, 14]

### Exemplo de fluxo com base da Receita Federal
```powershell
python leads_cnpj_receita.py --pasta ./cnpj_zips --uf PR --saida leads_pr_cnpj.csv
python consolida_leads.py --entrada leads_pr_cnpj.csv --saida leads_pr_consolidado.xlsx
python enriquece_leads.py --entrada leads_pr_consolidado_empresas.csv --prioridades A --limite 30
```[cite: 11, 13, 14]

### Exemplo via OpenStreetMap
```powershell
python leads_laboratorios.py --area "Curitiba" --nivel city --saida leads_cwb.csv
```[cite: 12]

## Boas Práticas e Privacidade

- **Não suba dados reais de leads, CNPJs, telefones, e-mails ou arquivos `.jsonl` de cache no GitHub.** Mantenha apenas os scripts e, se quiser, um exemplo pequeno com dados fictícios.
- Respeite limites de requisição, termos de uso de APIs externas e diretrizes de privacidade/LGPD no tratamento de dados de contato[cite: 12, 14].