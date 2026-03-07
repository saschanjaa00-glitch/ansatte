# Ansatte - bursdag og avdelingsverktøy

Webapp for å importere ansatte fra Excel og generere:

- bursdagskalender (`.ics`)
- liste over de som fyller 30, 40, 50, 57 eller 60 i år (`.csv`)
- utskriftsvennlig avdelingsliste i A4 Word-format med to kolonner (`.docx`)

All behandling skjer lokalt i nettleseren.

## Teknologi

- React + TypeScript + Vite
- `xlsx` for lesing av Excel
- `docx` for generering av Word-dokument

## Krav i data

Appen forventer kolonner som matcher (variasjoner i store/små bokstaver fungerer):

- navn
- telefon/mobil
- fødselsdato
- avdeling
- stillingskode/stillingskoder
- hovedstilling
- inkludert i skoleårsplanlegging

Filtrering:

- Kun rader der stillingskode (helst fra parentes) starter med `69`, `75`, `76` eller `79` blir med.
- Eksempel: `... (7960)` blir inkludert.

Avdeling:

- Hvis avdelingstekst inneholder `Avdelingsleder`, brukes teksten etter dette ordet.
- Eksempel: `St. Svithun vgs Avdelingsleder Realfag` blir `Realfag`.

## Kjøring lokalt

```bash
npm install
npm run dev
```

Prod build:

```bash
npm run build
```

## VS Code task

Det er laget en task som starter dev-server:

- `Run employee tool`
