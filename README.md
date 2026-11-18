<p align="center"><img src="src/LiveSlides.App/Assets/LiveSlides.png" width="160" alt="LiveSlides"></p>

<h1 align="center">LiveSlides</h1>

<p align="center">Manda in onda <b>presentazioni PowerPoint e PDF in sequenza, come un'unica presentazione</b>,<br>senza schermate nere, con un'immagine «tappo» tra una presentazione e l'altra.</p>

<p align="center">
  <a href="https://github.com/ncrdavide-svg/LiveSlides/releases/latest/download/LiveSlides-Setup.exe">
    <img src="docs/download-button.png" width="440" alt="⬇ Scarica LiveSlides — installer per Windows 10/11, ultima versione">
  </a>
</p>

<p align="center">
  <img src="https://img.shields.io/github/v/release/ncrdavide-svg/LiveSlides?label=versione&color=2ea44f" alt="versione"><br>
  Il pulsante scarica direttamente <code>LiveSlides-Setup.exe</code> · <a href="https://github.com/ncrdavide-svg/LiveSlides/releases">tutte le versioni e le novità</a>
</p>

> **Suite LIVE** · Con **[LiveSuite](https://github.com/ncrdavide-svg/LiveSuite)** installi e tieni aggiornate LiveCue, LiveSlides e LiveCount da un'unica piccola app (nessuna installazione): **[scarica LiveSuite.exe](https://github.com/ncrdavide-svg/LiveSuite/releases/latest/download/LiveSuite.exe)**.

---

## Come si installa

1. **Clicca il pulsante verde qui sopra**: parte subito il download di `LiveSlides-Setup.exe` (circa 55 MB), sempre l'ultima versione.
2. **Apri il file scaricato** (di solito è nella cartella *Download*).
3. Se compare la finestra blu **«Windows ha protetto il PC»**, clicca **Ulteriori informazioni** e poi **Esegui comunque**: succede perché l'installer non ha una firma digitale a pagamento, non perché sia pericoloso.
4. Segui l'installazione: puoi installare **per tutti gli utenti** (servono i permessi di amministratore) o **solo per te**.

Per aggiornare basta scaricare e installare la versione nuova sopra quella vecchia: impostazioni e scalette restano.

Per le presentazioni PowerPoint serve **PowerPoint** oppure **[LibreOffice](https://www.libreoffice.org/download/download-libreoffice/)** (gratuito); se mancano entrambi, l'installer lo segnala. I PDF e le immagini funzionano comunque. .NET è già incluso: non serve installare altro.

## Cosa fa

App per Windows della suite LIVE (con [LiveCue](https://github.com/ncrdavide-svg/LiveCue) e [LiveCount](https://github.com/ncrdavide-svg/LiveCount)).

- Le presentazioni sono riprodotte da **PowerPoint** (se installato) o da **LibreOffice**: animazioni, transizioni e filmati dentro le slide vanno come li ha progettati chi le ha fatte.
- Tre schermi, tipico del portatile: lo schermo del PC per l'operatore, uno per l'**uscita** a pieno schermo, uno per la **visualizzazione relatore** (slide attuale, successiva, note, tempo; con PDF e tappo ripete l'uscita).
- Avanti/indietro da tastiera, da telecomando per presentazioni o da remoto (Bitfocus Companion / Stream Deck, API HTTP sulla porta 8788).
- Scaletta con trascinamento, spunte per escludere un file, tappo all'inizio, tra le presentazioni e alla fine.
- **Salva show… / Apri show…** (Ctrl+S / Ctrl+O): lo show (cartella, ordine, spunte, tappo, dissolvenza) si salva in un file `.liveslides`. Salvato dentro la cartella delle presentazioni, funziona anche copiando la cartella su un altro PC. L'ordine si salva comunque da solo in `liveslides.json` nella cartella.

## Sviluppo

Requisiti: Windows 10/11, [.NET 9 SDK](https://dotnet.microsoft.com/download/dotnet/9.0).

```
dotnet run --project src/LiveSlides.App
dotnet test
```

### Installer per altri PC

```
powershell -ExecutionPolicy Bypass -File installer\build-installer.ps1 -Version 1.3.0
```

Crea in `artifacts\installer\` sia `LiveSlides-Setup-1.3.0.exe` sia `LiveSlides-Setup.exe` (stesso file, nome fisso) — serve [Inno Setup 6](https://jrsoftware.org/isinfo.php): `winget install JRSoftware.InnoSetup`.

Nella Release su GitHub vanno caricati **tutti e due**: il pulsante di download qui sopra punta a `releases/latest/download/LiveSlides-Setup.exe`, che funziona solo se l'ultima Release contiene il file con il nome fisso.

```
gh release create v1.3.0 artifacts\installer\LiveSlides-Setup-1.3.0.exe artifacts\installer\LiveSlides-Setup.exe --title "LiveSlides 1.3.0" --notes "..."
```

## Struttura

| Progetto | Contenuto |
|---|---|
| `src/LiveSlides.Core` | Logica senza interfaccia: la scaletta appiattita in un'unica sequenza di passi (`ShowSequence`), il file `liveslides.json` dello show. |
| `src/LiveSlides.App` | Interfaccia WPF, finestre di uscita e relatore, rendering dei PDF (motore PDF di Windows), riproduzione delle presentazioni con PowerPoint (COM) o LibreOffice (`Assets/lo_bridge.py`). |
| `tests/LiveSlides.Core.Tests` | Test della logica di `LiveSlides.Core`. |
| `installer/` | Script Inno Setup e script di build dell'installer. |
