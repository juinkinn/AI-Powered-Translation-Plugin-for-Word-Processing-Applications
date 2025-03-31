function onOpen() {
    DocumentApp.getUi().createMenu('AI Translator')
        .addItem('Open Sidebar', 'showSidebar')
        .addItem('Translate Selected Text', 'translateSelectedTextPrompt')
        .addItem('Translate Whole Document', 'translateWholeTextPrompt')
        .addToUi();
}

function showSidebar() {
    var html = HtmlService.createHtmlOutputFromFile('sidebar')
        .setTitle('AI Translator');
    DocumentApp.getUi().showSidebar(html);
}

function getSupportedLanguages() {
    var url = "https://libretranslate.com/languages";
    try {
        var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
        var json = JSON.parse(response.getContentText());
        
        if (!json || json.error) {
            Logger.log("Error fetching languages: " + (json ? json.error : "No data"));
            return [{ code: "auto", name: "Detect Language" }, { code: "en", name: "English" }];
        }

        var languages = json.map(function(lang) {
            return {
                code: lang.code,
                name: lang.name
            };
        });

        languages.unshift({ code: "auto", name: "Detect Language" });
        return languages;
    } catch (e) {
        Logger.log("Error: " + e.toString());
        return [{ code: "auto", name: "Detect Language" }, { code: "en", name: "English" }];
    }
}

function translateSelectedTextPrompt() {
    var ui = DocumentApp.getUi();
    var response = ui.prompt("Translate Selected Text", "Enter source language code, target language code, temperature, and style (e.g., en, es, 0.7, formal_academic):", ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() == ui.Button.OK) {
        var input = response.getResponseText().split(",");
        translateSelectedText(input[0].trim(), input[1].trim(), parseFloat(input[2].trim()), input[3] ? input[3].trim() : "neutral");
    }
}

function translateWholeTextPrompt() {
    var ui = DocumentApp.getUi();
    var response = ui.prompt("Translate Whole Document", "Enter source language code, target language code, temperature, and style (e.g., en, es, 0.7, professional_business):", ui.ButtonSet.OK_CANCEL);

    if (response.getSelectedButton() == ui.Button.OK) {
        var input = response.getResponseText().split(",");
        translateWholeText(input[0].trim(), input[1].trim(), parseFloat(input[2].trim()), input[3] ? input[3].trim() : "neutral");
    }
}

function translateSelectedText(sourceLang, targetLang, temperature, style) {
    var selection = DocumentApp.getActiveDocument().getSelection();
    if (!selection) {
        DocumentApp.getUi().alert("Error", "Please select some text.", DocumentApp.getUi().ButtonSet.OK);
        return;
    }

    var rangeElements = selection.getRangeElements();
    var translatedTexts = [];

    for (var i = 0; i < rangeElements.length; i++) {
        var element = rangeElements[i];
        var textElement = element.getElement().asText();
        var fullText = textElement.getText();
        var startOffset = element.getStartOffset();
        var endOffset = element.getEndOffsetInclusive();

        var selectedText = element.isPartial()
            ? fullText.substring(startOffset, endOffset + 1)
            : fullText;

        if (selectedText) {
            var translated = translateTextAPI(selectedText, sourceLang, targetLang, temperature, style);
            translatedTexts.push("Original: " + selectedText + "\nTranslated: " + translated);
        }
    }

    if (translatedTexts.length > 0) {
        DocumentApp.getUi().alert("Translation Result", translatedTexts.join("\n\n"), DocumentApp.getUi().ButtonSet.OK);
    }
}

function translateWholeText(sourceLang, targetLang, temperature, style) {
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var text = body.getText();
    if (!text) {
        DocumentApp.getUi().alert("Error", "Document is empty.", DocumentApp.getUi().ButtonSet.OK);
        return;
    }

    var translated = translateTextAPI(text, sourceLang, targetLang, temperature, style);
    DocumentApp.getUi().alert("Translation Result", translated, DocumentApp.getUi().ButtonSet.OK);
}

function translateTextAPI(text, sourceLang, targetLang, temperature, style) {
    var apiKey = ""; 
    var apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

    var stylePrompt = "";
    switch (style) {
        case "formal_academic":
            stylePrompt = "in a formal academic writing style suitable for scholarly papers";
            break;
        case "professional_business":
            stylePrompt = "in a professional business writing style suitable for corporate documents";
            break;
        case "technical":
            stylePrompt = "in a technical writing style suitable for manuals or scientific reports";
            break;
        case "casual":
            stylePrompt = "in a casual, conversational writing style";
            break;
        case "neutral":
        default:
            stylePrompt = "in a clear and neutral writing style";
            break;
    }

    var prompt = `Translate the following text from ${sourceLang === 'auto' ? 'auto' : sourceLang} to ${targetLang} ${stylePrompt}: ${text}, return only translated text.`;

    var payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt}
                ]
            }
        ],
        "generationConfig": {
            "temperature": parseFloat(temperature),
            "maxOutputTokens": 2000
        }
    };

    var options = {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
    };

    try {
        var response = UrlFetchApp.fetch(apiUrl, options);
        var json = JSON.parse(response.getContentText());
        if (json.error) {
            return "Error: " + json.error.message + " (Code: " + json.error.code + ")";
        }
        var result = json.candidates[0].content.parts[0].text;
        return result;
    } catch (e) {
        return "Error: " + e.toString();
    }
}