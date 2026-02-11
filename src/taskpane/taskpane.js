/* global document, Office, Word */
let isBookmarkInserted = false;

Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        console.log("Office is ready.");

        // --- FIX: Explicitly bind the button here ---
        const btn = document.getElementById("set-breakpoint-btn");
        const btnClear = document.getElementById("clear-breakpoint-btn");
        const btnGoTo = document.getElementById("go-to-breakpoint-btn");
        if (btn) {
            console.log("Button found, attaching listener.");
            btn.onclick = setBreakpoint;
            btnClear.setAttribute('disabled', 'disabled');
            btnGoTo.setAttribute('disabled', 'disabled');
            // Alternate method if onclick fails for some reason:
            // btn.addEventListener("click", setBreakpoint); 
        } else {
            console.error("ERROR: Could not find button with id 'set-breakpoint-btn'");
        }

        // Register the auto-update event
        Office.context.document.addHandlerAsync(
            Office.EventType.DocumentSelectionChanged,
            onSelectionChange
        );

                if (btnGoTo) {
            btnGoTo.onclick = goToBreakpoint;
        }

        // Initial update
        updateStats();
        document.getElementById("clear-breakpoint-btn").onclick = clearBreakpoint;
    }
});

async function goToBreakpoint() {
    await Word.run(async (context) => {
        try {
            // Cerca il segnalibro
            const bookmarkRange = context.document.getBookmarkRange("MyCharCounter_Breakpoint");
            
            // Il comando .select() sposta il cursore E scrolla la pagina fino al punto
            bookmarkRange.select();
            
            await context.sync();
        } catch (error) {
            console.log("Segnaposto non trovato o errore nello spostamento.");
        }
    });
}

function onSelectionChange(eventArgs) {
    updateStats();
}

async function clearBreakpoint() {
    await Word.run(async (context) => {
        try {
            const bookmarkRange = context.document.getBookmarkRange("MyCharCounter_Breakpoint");
            bookmarkRange.delete();
            await context.sync(); // Sync necessario per confermare l'eliminazione
            
            isBookmarkInserted = false;
            
            // Aggiornamento UI
            const btn = document.getElementById("set-breakpoint-btn");
            const btnClear = document.getElementById("clear-breakpoint-btn");
            const btnGoTo = document.getElementById("go-to-breakpoint-btn");

            btn.textContent = "Aggiungi segnaposto";
            btnClear.setAttribute('disabled', 'disabled');
            if (btnGoTo) btnGoTo.setAttribute('disabled', 'disabled');

        } catch (error) {
            console.log("Nothing to clear.");
        }
        updateStats();
    });
}

async function setBreakpoint() {
    await Word.run(async (context) => {
        const doc = context.document;
        const originalSelection = doc.getSelection();

        // 1. Controlli Preventivi sulla Selezione
        // Carichiamo il tipo di selezione per capire se è valida per inserire testo
        originalSelection.load("type");
        await context.sync();

        // Se l'utente ha selezionato un'immagine o una forma (InlineShape/Shape), 
        // l'inserimento di testo spesso fallisce o sostituisce l'immagine.
        // Se è "None", non c'è cursore.
        if (originalSelection.type === "None" || originalSelection.type === "InlineShape" || originalSelection.type === "Shape") {
            console.warn("Posizione non valida per il segnaposto (Immagine o Nessuna selezione).");
            // Opzionale: Mostra un avviso visibile all'utente
            return; 
        }

        // 2. Pulizia Vecchio Bookmark (Codice sicuro)
        try {
            const oldBookmark = doc.getBookmarkRangeOrNullObject("MyCharCounter_Breakpoint");
            await context.sync();
            if (!oldBookmark.isNullObject) {
                oldBookmark.delete();
                await context.sync();
            }
        } catch (e) {
            console.log("Errore rimozione vecchio (trascurabile): " + e);
        }

        // 3. Inserimento Nuovo (Protetto)
        try {
            // Ricarichiamo la selezione per sicurezza
            const selection = doc.getSelection();
            
            const colorSelect = document.getElementById("color");
            const selectedColor = colorSelect ? colorSelect.value : "red";
            const markerText = "🚩";

            // Inseriamo il testo
            const insertedRange = selection.insertText(markerText, Word.InsertLocation.replace);
            
            // Applichiamo proprietà
            insertedRange.font.color = selectedColor;
            insertedRange.insertBookmark("MyCharCounter_Breakpoint");

            // Spostiamo il cursore DOPO il segnaposto per evitare di scrivere "dentro" di esso
            insertedRange.getRange("After").select();

            await context.sync();

            // Aggiornamento UI solo se tutto è andato bene
            isBookmarkInserted = true;
            const btn = document.getElementById("set-breakpoint-btn");
            const btnClear = document.getElementById("clear-breakpoint-btn");
            const btnGoTo = document.getElementById("go-to-breakpoint-btn");
            const btnPrint = document.getElementById("prepare-print-btn");

            if (btnClear) btnClear.removeAttribute("disabled");
            if (btnGoTo) btnGoTo.removeAttribute("disabled");
            if (btnPrint) btnPrint.removeAttribute("disabled");
            if (btn) btn.textContent = "Aggiorna segnaposto";

            updateStats();

        } catch (error) {
            console.error("ERRORE CRITICO durante l'inserimento: ", error);
            // Se fallisce qui (GeneralException), è perché Word si rifiuta di scrivere in quel punto specifico.
            // Non possiamo farci molto se non avvisare.
        }
    });
}



async function updateStats() {
    await Word.run(async (context) => {
        const doc = context.document;
        const currentSelection = doc.getSelection();
        const bodyRange = doc.body;

        // Prepariamo i caricamenti
        bodyRange.load("text");
        currentSelection.load("text");

        // Calcolo posizione cursore
        // Nota: Questo è il punto critico che fallisce nei link
        const cursorRange = currentSelection.getRange("Start");
        const startToCursorRange = cursorRange.expandTo(bodyRange.getRange("Start"));
        startToCursorRange.load("text");

        // Variabili per i risultati
        let totalCount = 0;
        let selectionCount = 0;
        let beforeCursorCount = 0;
        let breakpointCount = "N/A";

        // --- BLOCCO 1: Dati Standard ---
        try {
            // Proviamo a sincronizzare per ottenere i dati base
            await context.sync();

            totalCount = bodyRange.text.length;
            selectionCount = currentSelection.text.length;
            beforeCursorCount = startToCursorRange.text.length;

        } catch (error) {
            // Se fallisce qui (es. dentro un link strano), usciamo silenziosamente
            // senza aggiornare l'UI con dati errati, ma senza rompere l'addon.
            console.log("Salto aggiornamento stat: selezione instabile o in campo speciale.");
            return; 
        }

        // --- BLOCCO 2: Dati Segnaposto ---
        // Questo blocco era già protetto, ma lo manteniamo separato
        try {
            const bookmarkRange = doc.getBookmarkRange("MyCharCounter_Breakpoint");
            const bookmarkEnd = bookmarkRange.getRange("End");
            
            // Creiamo il range dal segnaposto al cursore attuale
            const measuredRange = bookmarkEnd.expandTo(currentSelection.getRange("Start"));
            measuredRange.load("text");
            
            await context.sync();
            
            breakpointCount = measuredRange.text.length.toString();
        } catch (error) {
            // Se il segnaposto non esiste o il calcolo fallisce, resta "N/A"
            // Ignoriamo l'errore ItemNotFound qui perché è normale se non c'è il bookmark
        }

        // Aggiorna UI solo se siamo arrivati vivi fin qui
        updateUI(totalCount, selectionCount, beforeCursorCount, breakpointCount);
    });
}

function updateUI(total, selected, before, breakpoint) {
    document.getElementById("total-count").innerText = total;
    document.getElementById("selection-count").innerText = selected;
    document.getElementById("cursor-count").innerText = before;
    document.getElementById("breakpoint-count").innerText = breakpoint;
}
