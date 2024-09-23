Office.onReady(function(info) {
    if (info.host === Office.HostType.PowerPoint) {
        document.getElementById("fetchData").onclick = fetchDataFromGroqAndOpenAI;
    }
});

async function fetchDataFromGroqAndOpenAI() {
    const resultDiv = document.getElementById("result");
    resultDiv.innerHTML = "Fetching data...";

    try {
        // Fetch data from Groq AI
        const groqResponse = await fetch("https://api.groq.com/v1/your-endpoint", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer YOUR_GROQ_API_KEY"
            },
            body: JSON.stringify({
                messages: [
                    { role: "user", content: "Explain the importance of low latency LLMs" }
                ],
                model: "llama3-8b-8192"
            })
        });

        const groqData = await groqResponse.json();
        const groqContent = groqData.choices[0].message.content;

        // Fetch perplexity data from OpenAI
        const openaiResponse = await fetch("https://api.openai.com/v1/engines/davinci-codex/completions", {
            method: "POST",
            headers: {
                "Content-Type": "application/json",
                "Authorization": "Bearer YOUR_OPENAI_API_KEY"
            },
            body: JSON.stringify({
                prompt: "Calculate the perplexity of the following text: " + groqContent,
                max_tokens: 50
            })
        });

        const openaiData = await openaiResponse.json();
        const openaiContent = openaiData.choices[0].text;

        resultDiv.innerHTML = `<h2>Groq AI Response:</h2><p>${groqContent}</p><h2>OpenAI Perplexity:</h2><p>${openaiContent}</p>`;
    } catch (error) {
        resultDiv.innerHTML = "Error fetching data: " + error.message;
    }
}
