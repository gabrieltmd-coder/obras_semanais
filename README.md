# Registro Semanal de Obras — Rumo Logística

## Como rodar

```bash
pip install -r requirements.txt
python app.py
```

Acesse: http://localhost:5000

## Rotas
- `/` — listagem com filtros
- `/novo` — cadastrar registro
- `/editar/<id>` — editar registro
- `/export/excel` — download Excel para Power BI

## Estrutura
```
obras_semanais/
├── app.py              # Backend Flask
├── templates/
│   ├── base.html       # Layout Rumo Logística
│   ├── index.html      # Listagem
│   └── form.html       # Cadastro/edição
├── data/
│   └── registros.json  # Persistência
├── exports/            # Excel gerado
└── requirements.txt
```
