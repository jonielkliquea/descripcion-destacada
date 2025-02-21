import OpenAI from "openai";
import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";

const filePath = path.resolve("src", "Book.xlsx");
const fileBuffer = fs.readFileSync(filePath);
const workbook = XLSX.read(fileBuffer, { type: "buffer" });

const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet);

const systemPrompt = `Eres un asistente que solo responde con un JSON estructurado sin ningún tipo de explicación o texto adicional.
Debes convertir la descripción del producto en una lista de puntos clave en un array "bullet" y generar código HTML en "code".  
El formato de salida debe ser exactamente este:

{
  "Name":"nombre dado sin cambios", 
  "ProductDescription":"Mejorarlo tomando en cuenta la data suministrada",
  "bullet": [
    "Descripción corta del punto 1",
    "Descripción corta del punto 2",
    "Descripción corta del punto 3"
  ],
  "code": "<body>...</body>"
}

Reglas estrictas:
- Lo dado en "Name" no cambia.
- El "ProductDescription" debe ser mejorado basado en la data suministrada.
- Si la descripción está en HTML, conviértela en texto plano.
- "bullet" debe contener al menos 3 elementos.
- "code" debe ser una lista <ul> en HTML.
- No devuelvas explicaciones, solo el JSON.

Ejemplo de respuesta correcta:

{
  "Name":"nombre dado sin cambios",
  "ProductDescription":"Mejorarlo tomando en cuenta la data suministrada",
  "bullet": [
    "Silueta ajustada que realza las curvas",
    "Escote cruzado en el frente con canal",
    "Cargaderas ajustables para mejor ajuste"
  ],
  "code": "<body><h1>Descripción del Producto</h1><ul><li><strong>Silueta ajustada:</strong> Realza las curvas.</li><li><strong>Escote cruzado:</strong> Agrega sofisticación.</li><li><strong>Cargaderas ajustables:</strong> Para mayor comodidad.</li></ul></body>"
}

IMPORTANTE: Si la descripción no tiene suficiente información, genera al menos 3 puntos clave basándote en la información proporcionada.  
No devuelvas ningún texto fuera del JSON.`;

const openai = new OpenAI({
  baseURL: "http://localhost:1234/v1/",
  apiKey: "sk-cb409945a45d495d9310f7ccba0b33f9",
});

const results = [];

async function processExcelData() {
  for (const [index, row] of data.entries()) {
    const userPrompt = JSON.stringify(row, null, 2);
    try {
      const completion = await openai.chat.completions.create({
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userPrompt },
        ],
        model: "hermes-3-llama-3.2-3b",
        temperature: 0.7,
        max_tokens: 5000,
        stream: false,
      });

      const response = completion.choices[0].message.content;

      if (!response || !response.startsWith("{")) {
        console.error(`⚠️ Respuesta no válida en fila ${index + 1}:`, response);
        continue;
      }

      const jsonResponse = JSON.parse(response);

      results.push({
        Name: String(jsonResponse["Name"]), 
        ProductDescription: String(jsonResponse["ProductDescription"]),
        Bullet: jsonResponse["bullet"].join("\n"),
        Code: jsonResponse["code"],
      });

      console.log(`✅ Procesado: ${index + 1} / ${data.length} (${((index + 1) / data.length * 100).toFixed(2)}%)`);
    } catch (error) {
      console.error(` Error en fila ${index + 1}:`, error);
    }
  }
  saveToExcel(results);
}

function saveToExcel(data) {
  const newWorkbook = XLSX.utils.book_new();
  const newWorksheet = XLSX.utils.json_to_sheet(data);
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Resultados");

  const outputFilePath = path.resolve("src", "output.xlsx");
  XLSX.writeFile(newWorkbook, outputFilePath);
  console.log("✅ Archivo guardado en:", outputFilePath);
}

processExcelData();