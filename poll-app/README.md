# Poll live MPS

Aplicație independentă de poll live, fără cont Claude. Publicul votează de pe telefon, iar prezentatorul pornește și oprește întrebările din `/master`.

## Pagini

- `/` vedere public: așteptare, întrebare deschisă, apoi graficul de rezultate
- `/master` panou prezentator (cere PIN): Start, Stop + arată rezultate, ștergere voturi
- `/join` ecran mare cu codul QR către pagina publică

## Variabile de mediu

- `MASTER_PIN` (obligatoriu): PIN-ul pentru `/master`

## Deploy pe Netlify (recomandat, fără stocare externă)

1. Netlify, Add new site, Import from Git, alege repo-ul `prezentare_mps`.
2. Base directory: `poll-app`. Publish directory: `public`. Build command: gol.
3. Site configuration, Environment variables: adaugă `MASTER_PIN`.
4. Deploy. Voturile se păstrează automat în Netlify Blobs.

## Deploy pe Vercel

1. Vercel, Add New Project, importă repo-ul, Root Directory: `poll-app`. Framework: Other.
2. Storage, Marketplace, adaugă Upstash Redis (Vercel KV) și conectează-l la proiect. Variabilele `KV_REST_API_URL` și `KV_REST_API_TOKEN` se adaugă automat.
3. Adaugă `MASTER_PIN` la Environment Variables și redeploy.

Fără stocare configurată pe Vercel, serverul folosește memorie temporară, iar panoul afișează un avertisment.

## Rulare locală

```
npm install
MASTER_PIN=1234 npm run dev
npm test
```

Deschide `http://localhost:3000/master` și `http://localhost:3000/`.

## Întrebări

Se editează în `lib/polls.js`.
