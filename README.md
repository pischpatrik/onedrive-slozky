# OneDrive Slozky PWA

Mobilni PWA pro rychly pristup k vybranym slozkam z OneDrive.

## Co uz umi

- prihlaseni k OneDrive pres Microsoft ucet
- seznam oblibenych slozek
- prochazeni podslozek a souboru
- otevreni beznych souboru
- hledani vhodne appky pro mene bezne koncovky
- instalace na plochu jako PWA
- demo rezim i skutecny OneDrive rezim
- lokalni nastaveni primo v aplikaci

## Jak to spustit lokalne

1. Otevri terminal ve slozce projektu.
2. Spust jednoduchy server:

```powershell
python -m http.server 4174
```

3. Otevri `http://127.0.0.1:4174`.
4. V aplikaci klikni na `Nastaveni`.

## Jak nastavit skutecny OneDrive

### 1. Vytvor Microsoft aplikaci

V Microsoft Entra admin centru:

1. Otevri `App registrations`.
2. Klikni `New registration`.
3. Vypln nazev, treba `OneDrive Slozky`.
4. U `Supported account types` zvol:
   `Accounts in any organizational directory and personal Microsoft accounts`.
5. U `Redirect URI` vyber typ `Single-page application`.
6. Pridej presnou adresu, kterou vidi aplikace v `Nastaveni`.

Pro lokalni vyvoj to bude typicky:

- `http://127.0.0.1:4174/`

Pro GitHub Pages to bude typicky:

- `https://TVE-JMENO.github.io/NAZEV-REPA/`

Pak uloz a zkopiruj `Application (client) ID`.

### 2. Pridej opravneni

V `API permissions` nech nebo pridej delegovana opravneni:

- `Files.Read`
- `User.Read`
- `offline_access`
- `openid`
- `profile`

### 3. Vloz client ID do aplikace

1. Otevri PWA.
2. Klikni `Nastaveni`.
3. Vloz `Microsoft client ID`.
4. Tenant nech jako `common`.
5. Uloz nastaveni.
6. Klikni `Prihlasit k OneDrive`.

Tohle delas jen ty jako autor aplikace. Bezny uzivatel uz nic z Entra resit nebude, jen klikne na `Prihlasit k OneDrive`.

## Jak to dat na GitHub Pages

1. Vytvor nove repo na GitHubu.
2. Nahraj do nej vsechny soubory z tehle slozky.
3. Na GitHubu otevri `Settings -> Pages`.
4. U `Build and deployment` nastav:
   `Deploy from a branch`
5. Vyber branch `main` a slozku `/(root)`.
6. Uloz.
7. GitHub ti vytvori adresu ve tvaru:
   `https://TVE-JMENO.github.io/NAZEV-REPA/`
8. Tu samou adresu pridej i do Microsoft Entra jako dalsi `Redirect URI`.
9. Otevri tuto adresu, v aplikaci dej `Nastaveni`, vloz client ID a prihlas se.

Pokud chces, aby bezni uzivatele vubec nevideli `Nastaveni`, nech v `config.js`:

```js
allowRuntimeSettings: false
```

Po vlozeni tveho `microsoftClientId` bude aplikace pro ostatni zobrazovat uz jen prihlaseni.

Soubor `.nojekyll` uz je pripraveny, aby GitHub Pages nic zbytecne neupravoval.

## Jak z toho udelat appku v mobilu

### iPhone

1. Otevri GitHub Pages adresu v `Safari`.
2. Klikni `Sdilet`.
3. Zvol `Na plochu`.

### Android

1. Otevri GitHub Pages adresu v `Chrome`.
2. Zvol `Install app` nebo `Pridat na plochu`.

## Poznamka

Token je ulozeny v prohlizeci, aby aplikace zustala jednoducha a staticka. Pro plne produkcni nasazeni je pozdeji vhodne zvazit robustnejsi auth vrstvu nebo knihovnu `MSAL`.
