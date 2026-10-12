# Míchání studentských dvojic (VND)

Jednoduchá, moderní webová aplikace pro náhodné sestavování dvojic (a trojice v případě lichého počtu studentů) pro studijní skupiny.

Aplikace běží kompletně v prohlížeči (frontend bez nutnosti backendového serveru) a je připravena pro přímé nasazení na **GitHub Pages**.

---

## 🚀 Hlavní funkce

- **Automatické načtení z repozitáře**: Aplikace při startu automaticky načte soubor `VND26_studenti.xlsx`.
- **Nahrání vlastního souboru**: Možnost přetáhnout (Drag & Drop) nebo vybrat jakýkoliv jiný `.xlsx` soubor studentů ve stejné struktuře:
  1. sloupec: *Křestní jméno*
  2. sloupec: *Příjmení*
  3. sloupec: *Číslo skupiny* (např. `42` nebo `B261-11VND studijní skupina 42`)
- **Inteligentní roletka**: Automaticky nabídne všechny dostupné skupiny s počtem studentů v závorce.
- **Správa docházky (chybějící studenti)**:
  - V seznamu skupiny lze kliknutím na studenta nebo tlačítko `❌ Chybí` označit nepřítomné.
  - Aplikace automaticky přepočítá počet dvojic/trojic podle skutečně přítomných studentů.
  - Kdykoliv lze chybějící studenty vrátit jednotlivě nebo hromadně tlačítkem `Vrátit všechny chybějící`.
- **Náhodné míchání (Fisher-Yates)**:
  - Zamíchá pouze přítomné studenty.
  - Při sudém počtu: čistě dvojice.
  - Při lichém počtu: dvojice + přesně **1 trojice** (nikdo nezůstane sám).
- **Ruční výměna studentů kliknutím**:
  - Po namíchání stačí kliknout na jednoho studenta (zvýrazní se žlutě) a poté na druhého – studenti se okamžitě prohodí.
  - Funguje mezi dvojicemi i mezi dvojicí a trojicí.
  - TXT výstup i náhled se okamžitě automaticky aktualizují.
- **Export do TXT**:
  - Tlačítko **Uložit do TXT** stáhne soubor `skupina_XX_dvojice.txt`.
  - Formát: **jedna dvojice/trojice na jeden řádek**.
  - Opatřeno UTF-8 BOM pro bezchybné zobrazení české diakritiky v Poznámkovém bloku (Notepad) ve Windows i Excelu.
- **Kopírování do schránky**: Jedním kliknutím zkopíruje vygenerovaný seznam.

---

## 💻 Spuštění na počítači (lokálně)

1. Stačí otevřít soubor `index.html` v libovolném moderním prohlížeči (Chrome, Edge, Firefox, Safari).
2. Volitelně můžete spustit lokální HTTP server (např. přes Python v příkazovém řádku):
   ```bash
   python -m http.server 8000
   ```
   A otevřít v prohlížeči adresu `http://localhost:8000`.

---

## 🌐 Publikace na GitHub Pages

1. **Vytvořte nový repozitář na GitHubu** (např. `AntigravityMichani` nebo `vnd-dvojice`).
2. **Nahrajte soubory do repozitáře**:
   ```bash
   git init
   git add .
   git commit -m "Initial commit - aplikace na michani dvojic"
   git branch -M main
   git remote add origin https://github.com/VASE-JMENO/NAZEV-REPOZITARE.git
   git push -u origin main
   ```
3. **Zapněte GitHub Pages**:
   - V repozitáři na GitHubu přejděte do **Settings** -> **Pages**.
   - V sekci **Build and deployment**:
     - *Source*: **Deploy from a branch**
     - *Branch*: **main** / root (`/`)
   - Klikněte na **Save**.
4. Během 1–2 minut bude vaše aplikace dostupná online na adrese:  
   `https://VASE-JMENO.github.io/NAZEV-REPOZITARE/`
