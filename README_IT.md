[! [NPM Deployment Status] (https://travis-ci.org/OfficeDev/office-js.svg?branch=release)] (https://travis-ci.org/OfficeDev/office-js/builds)

# API JavaScript di Office

L'API JavaScript per Office consente di creare applicazioni Web che interagiscono con i modelli di oggetti nelle applicazioni host di Office. La tua applicazione farà riferimento alla libreria office.js, che è un caricatore di script. La libreria office.js carica i modelli di oggetti applicabili all'applicazione Office che esegue il componente aggiuntivo.

<br />

## Informazioni sul pacchetto NPM

Il pacchetto NPM per Office.js è una copia di ciò che viene pubblicato sul CDN ufficiale di Office.js "evergreen", in ** <https://appsforoffice.microsoft.com/lib/1/hosted/office.js> * *.

Mentre la CDN di Office.js contiene tutte le API di Office.js attualmente disponibili in qualsiasi momento, ciascuna versione del pacchetto NPM per Office.js contiene solo le API di Office.js che erano disponibili nel momento in cui tale versione del Il pacchetto NPM è stato creato.

### Scenari di destinazione

Il pacchetto NPM per Office.js è inteso come un modo per ottenere la tua copia (non CDN) dei file Office.js, che puoi quindi servire staticamente dal tuo sito invece di utilizzare la CDN. Questo pacchetto NPM viene principalmente fornito per affrontare i seguenti scenari:

1. Se si sta sviluppando un componente aggiuntivo dietro un firewall, non è possibile accedere al CDN di Office.js.

2. Se è necessario l'accesso offline alle API di Office.js (ad esempio, per facilitare il debug offline).

### Migliori pratiche

Le migliori pratiche per l'utilizzo del pacchetto NPM di Office.js includono:

- Aggiorna periodicamente il tuo pacchetto NPM (per accedere a nuove API e / o correzioni di bug che potrebbero non essere disponibili nella versione corrente del pacchetto).

- Usa il pacchetto NPM secondo le istruzioni in [Uso del pacchetto NPM] (# using-the-npm-package); non tentare di importare il pacchetto NPM come si potrebbe comunemente fare con altri pacchetti NPM.

- Non utilizzare il pacchetto NPM in un componente aggiuntivo inviato per la pubblicazione su [AppSource] (https://appsource.microsoft.com/marketplace/apps?product=office). I componenti aggiuntivi pubblicati su AppSource devono utilizzare il CDN di Office.js.

- Utilizzare le definizioni TypeScript per Office.js come descritto in [Definizioni IntelliSense] (# definizioni intellisense).

<br />

## Installazione del pacchetto NPM

Per installare "office-js" localmente tramite il pacchetto NPM, eseguire il seguente comando:

    npm install @ microsoft / office-js --save

<br />

## Utilizzo del pacchetto NPM

L'installazione del pacchetto NPM localmente crea una serie di file statici di Office.js nella cartella `node_modules \ @microsoft \ office-js \ dist` della directory in cui è stato eseguito il comando` npm install`. Per utilizzare il pacchetto NPM, effettuare le seguenti operazioni:

1. Sia manualmente o come parte di uno script di compilazione (ad esempio, `CopyWebpackPlugin` se si sta utilizzando Webpack), i file vengono serviti da una destinazione di propria scelta (ad esempio, da` / assets / office-js / ` directory del tuo server web).

2. Riferimento che lo
