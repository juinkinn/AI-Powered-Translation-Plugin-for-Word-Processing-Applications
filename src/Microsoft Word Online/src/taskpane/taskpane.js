Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
      console.log("Office.js initialized for Word");
      populateLanguages();
  } else {
      console.error("Not running in Word:", info.host);
  }
});

function populateLanguages() {
  fetch("https://libretranslate.com/languages")
      .then(response => response.json())
      .then(languages => {
          const sourceLangSelect = document.getElementById("sourceLang");
          const targetLangSelect = document.getElementById("targetLang");

          languages.forEach(lang => {
              const option1 = document.createElement("option");
              option1.value = lang.code;
              option1.text = `${lang.name} (${lang.code})`;
              sourceLangSelect.appendChild(option1);

              if (lang.code !== "auto") {
                  const option2 = document.createElement("option");
                  option2.value = lang.code;
                  option2.text = `${lang.name} (${lang.code})`;
                  targetLangSelect.appendChild(option2);
              }
          });

          sourceLangSelect.value = "auto";
          targetLangSelect.value = "en";
      })
      .catch(error => {
          console.error("Error fetching languages:", error);
          const fallback = [
              { code: "auto", name: "Detect Language" },
              { code: "en", name: "English" }
          ];
          fallback.forEach(lang => {
              const option = document.createElement("option");
              option.value = lang.code;
              option.text = `${lang.name} (${lang.code})`;
              document.getElementById("sourceLang").appendChild(option);
              if (lang.code !== "auto") {
                  document.getElementById("targetLang").appendChild(option.cloneNode(true));
              }
          });
          document.getElementById("sourceLang").value = "auto";
          document.getElementById("targetLang").value = "en";
      });
}

function updateSliderValue(value) {
  document.getElementById("tempValue").textContent = value;
}

function showPopup(content) {
  document.getElementById("popupContent").textContent = content;
  document.getElementById("customPopup").style.display = "flex"; 
}

function closePopup() {
  document.getElementById("customPopup").style.display = "none"; 
}

async function translateSelectedText() {
  console.log("translateSelectedText called");
  const sourceLang = document.getElementById("sourceLang").value;
  const targetLang = document.getElementById("targetLang").value;
  const temperature = document.getElementById("temperature").value;
  const style = document.getElementById("style").value;

  try {
      await Word.run(async (context) => {
          const selection = context.document.getSelection();
          selection.load("text");
          await context.sync();

          if (!selection.text) {
              showPopup("Please select some text.");
              return;
          }

          const translated = await translateTextAPI(selection.text, sourceLang, targetLang, temperature, style);
          showPopup(translated); 
      });
  } catch (error) {
      console.error("Error in translateSelectedText:", error);
      showPopup(`Error: ${error.message}`);
  }
}

async function translateWholeText() {
  console.log("translateWholeText called");
  const sourceLang = document.getElementById("sourceLang").value;
  const targetLang = document.getElementById("targetLang").value;
  const temperature = document.getElementById("temperature").value;
  const style = document.getElementById("style").value;

  try {
      await Word.run(async (context) => {
          const body = context.document.body;
          body.load("text");
          await context.sync();

          if (!body.text) {
              showPopup("Document is empty.");
              return;
          }

          const translated = await translateTextAPI(body.text, sourceLang, targetLang, temperature, style);
          showPopup(translated);
      });
  } catch (error) {
      console.error("Error in translateWholeText:", error);
      showPopup(`Error: ${error.message}`);
  }
}

async function translateTextAPI(text, sourceLang, targetLang, temperature, style) {
  const apiKey = ""; 
  const apiUrl = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${apiKey}`;

  const stylePrompt = {
      "formal_academic": "in a formal academic writing style suitable for scholarly papers",
      "professional_business": "in a professional business writing style suitable for corporate documents",
      "technical": "in a technical writing style suitable for manuals or scientific reports",
      "casual": "in a casual, conversational writing style",
      "neutral": "in a clear and neutral writing style"
  }[style] || "in a clear and neutral writing style";

  const prompt = `Translate the following text from ${sourceLang === 'auto' ? 'auto' : sourceLang} to ${targetLang} ${stylePrompt}: ${text}, return only translated text.`;

  const payload = {
      contents: [{ parts: [{ text: prompt }] }],
      generationConfig: { temperature: parseFloat(temperature), maxOutputTokens: 2000 }
  };

  try {
      const response = await fetch(apiUrl, {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify(payload)
      });
      const json = await response.json();
      if (json.error) {
          return `Error: ${json.error.message} (Code: ${json.error.code})`;
      }
      return json.candidates[0].content.parts[0].text;
  } catch (e) {
      return `Error: ${e.toString()}`;
  }
}