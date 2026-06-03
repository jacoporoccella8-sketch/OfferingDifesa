# Defence Market Daily Scraper - NTT DATA Italia

Market analysis automatizzata sul segmento **Difesa italiano**, commissionata da
NTT DATA Italia. Ogni giorno lo script interroga l'API Anthropic (con web search
integrato), raccoglie dati sui player monitorati, genera un report Excel a tre
fogli e invia una mail HTML riassuntiva con l'Excel in allegato.

## Cosa raccoglie per ogni player

- Fatturato piu recente
- Stima della spesa IT annua, con fonte/metodologia
- Servizi IT acquistati o di interesse
- Piano strategico digitale / di trasformazione IT
- Storico gare ANAC (CIG, importo, aggiudicatario, anno)
- Trend futuri
- Note utili per NTT DATA

## Player monitorati (11)

| # | Player | Mercato | Complessita procurement |
|---|--------|---------|--------------------------|
| 1 | Ministero della Difesa | Pubblico | Molto alta |
| 2 | Forze Armate (Esercito/Marina/AM) | Pubblico | Alta |
| 3 | Difesa Servizi | Pubblico | Media |
| 4 | Segretariato Generale Difesa/DNA | Pubblico | Alta |
| 5 | Leonardo | Privato | Alta |
| 6 | Fincantieri | Privato | Alta |
| 7 | MBDA Italia | Privato | Alta |
| 8 | Elettronica Group | Privato | Alta |
| 9 | Thales Alenia Space Italia | Privato | Alta |
| 10 | Avio Aero | Privato | Alta |
| 11 | Leonardo DRS | Privato | Alta |

## Scaglionamento a rotazione (vincolo Tier 1)

L'account Anthropic e su **Tier 1 con budget molto basso (circa 5 USD)**, quindi
i rate limit sono stringenti. Interrogare tutti gli 11 player in un singolo run
fa scattare l'errore **429 (rate_limit_error)** gia dopo il primo player.

La soluzione adottata e l'opposta della parallelizzazione: **rallentare e rendere
lo script robusto**. Nello specifico:

- **3 player per esecuzione, a rotazione.** Il parametro `PLAYERS_PER_RUN`
  (default `3`) definisce quanti player vengono processati per run. Lo stato
  della rotazione (indice del prossimo player) e salvato in
  `data/rotation_state.json`.
  - Giorno 1: player 1-3
  - Giorno 2: player 4-6
  - Giorno 3: player 7-9
  - Giorno 4: player 10-11 (+ ricomincia dal player 1)
  - Poi il ciclo riparte aggiornando i dati.
  - In **4 giorni** vengono coperti tutti gli 11 player, poi si cicla.
- **I player non processati in un dato giorno** mantengono nell'Excel e nella
  mail **l'ultimo dato valido** salvato in `data/history.json`. Cosi ogni report
  contiene comunque il quadro completo di tutti gli 11 player.
- **Pausa di 10 secondi** tra un player e l'altro (`PAUSE_BETWEEN_PLAYERS`).
- **Retry con backoff esponenziale sul 429**: attese 20s, 40s, 80s, max 3 retry
  (`MAX_RETRIES`, `BACKOFF_SECONDS`). Solo dopo l'ultimo retry fallito il player
  viene segnato come errore.
- **Web search bilanciato**: massimo 4 ricerche per player
  (`WEB_SEARCH_MAX_USES`), compromesso tra qualita e consumo quota.
- **Logging dell'errore vero**: per ogni player viene loggato l'esito
  (successo, oppure 429 rate limit, timeout, JSON non parsato, ecc.) con livelli
  INFO ed ERROR, cosi dal log si capisce esattamente perche un player e fallito.
- **Guardia sull'invio mail**: se ZERO player producono dati validi (neppure
  storici), la mail NON viene inviata; lo script logga l'errore ed esce con
  codice diverso da zero, cosi il run risulta rosso su GitHub Actions.

### Come aumentare i player per run (tier API superiore)

Se si passa a un tier API superiore con rate limit piu generosi, basta **alzare
il parametro `PLAYERS_PER_RUN`** in cima a `defence_market_scraper.py`. Ad
esempio `PLAYERS_PER_RUN = 11` processa tutti i player in un solo run. Si possono
inoltre ridurre `PAUSE_BETWEEN_PLAYERS` e aumentare `WEB_SEARCH_MAX_USES` per
ottenere dati piu ricchi.

## Output Excel (3 fogli)

1. **Dashboard** - tabella principale con tutti gli 11 player. Colonne:
   Player, Mercato, Complessita Procurement, Fatturato, Spesa IT Stimata,
   Fonte Stima, Servizi IT Acquistati, Piano Strategico IT, N. Gare Trovate,
   Trend Futuro, Delta vs Ieri, Note per NTT DATA.
2. **Storico Gare** - tutte le gare trovate (Player, CIG, Oggetto, Importo,
   Aggiudicatario, Anno).
3. **Dettaglio Player** - una scheda per ogni player con tutti i campi.

## Mail

Mail HTML inviata da un account Gmail (SMTP `smtp.gmail.com`, porta 587, STARTTLS)
a `jacopo.roccella@nttdata.com`. Contiene 4 KPI box (player monitorati, con
stima spesa IT, gare trovate, variazioni vs ieri), la tabella riassuntiva dei
player e l'Excel in allegato.

> **Gmail**: usare una **App Password** (non la password normale dell'account),
> con verifica in due passaggi attiva. Inserirla nel secret `SMTP_PASS`.

## Delta vs giorno precedente

Per ogni player processato lo script confronta i dati nuovi con quelli salvati
in `data/history.json` e produce un testo che riassume le variazioni su spesa
IT, fatturato, trend e numero di gare. Lo storico viene poi aggiornato.

## Esecuzione schedulata (GitHub Actions)

Il workflow `.github/workflows/daily_scraper.yml` gira ogni giorno via cron e
puo essere lanciato manualmente (`workflow_dispatch`).

> **Nota sull'orario**: il cron e `30 6 * * *` (06:30 UTC), che corrisponde alle
> **07:30 CET** in ora solare. Durante l'ora legale (CEST) le 06:30 UTC sono le
> 08:30 locali; GitHub Actions supporta solo UTC, quindi per avere sempre le
> 07:30 locali esatte occorrerebbe alternare il cron tra stagioni.

Il workflow ha `permissions: contents: write` e, a fine run, fa **commit
automatico su branch `main`** di `data/history.json` e `data/rotation_state.json`.

## Variabili d'ambiente / GitHub Secrets

| Secret | Descrizione |
|--------|-------------|
| `ANTHROPIC_API_KEY` | Chiave API Anthropic |
| `SMTP_HOST` | `smtp.gmail.com` |
| `SMTP_PORT` | `587` |
| `SMTP_USER` | Account Gmail mittente |
| `SMTP_PASS` | App Password Gmail |
| `RECIPIENT_EMAIL` | `jacopo.roccella@nttdata.com` |

## Modello Anthropic

Lo script usa `claude-sonnet-4-6` (versione datata `claude-sonnet-4-6-20260218`).
**Non** usare `claude-sonnet-4-20250514` (deprecato, errore 404).

## Esecuzione locale

```bash
pip install -r requirements.txt
cp .env.example .env   # e valorizzare le variabili
python defence_market_scraper.py
```

## Struttura del progetto

```
defence-market-scraper/
  defence_market_scraper.py
  requirements.txt
  README.md
  .env.example
  data/.gitkeep
  output/.gitkeep
  .github/workflows/daily_scraper.yml
```
