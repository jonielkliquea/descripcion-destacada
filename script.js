import OpenAI from "openai";
import fs from "fs";
import path from "path";
import * as XLSX from "xlsx";

const filePath = path.resolve("src", "Booktt.xlsx");
const fileBuffer = fs.readFileSync(filePath);
const workbook = XLSX.read(fileBuffer, { type: "buffer" });

const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const data = XLSX.utils.sheet_to_json(worksheet); 

const systemPrompt = `Eres un asistente que solo responde con un JSON estructurado sin ningún tipo de explicación o texto adicional.

Debes convertir la descripción del producto en una lista de puntos clave en un array "bullet" y generar código HTML en "code".  
El formato de salida debe ser exactamente este:

{
  "bullet": [
    "Descripción corta del punto 1",
    "Descripción corta del punto 2",
    "Descripción corta del punto 3"
  ],
  "code": "<body>...</body>"
}

Reglas estrictas:
- No agregues explicaciones ni contexto, **solo responde con el JSON**.
- No modifiques el formato de salida.
- Asegúrate de que "bullet" sea un array de al menos 3 elementos.
- Asegúrate de que "code" sea un bloque HTML con una lista <ul>.
- No incluyas comentarios dentro del JSON.

Ejemplo de respuesta correcta:

{
  "bullet": [
    "Silueta ajustada que realza las curvas",
    "Escote cruzado en el frente con canal",
    "Cargaderas ajustables para mejor ajuste"
  ],
  "code": "<body><h1>Descripción del Producto (este texto no cambia) </h1><ul><li><strong>Silueta ajustada:</strong> Realza las curvas.</li><li><strong>Escote cruzado:</strong> Agrega sofisticación.</li><li><strong>Cargaderas ajustables:</strong> Para mayor comodidad.</li></ul></body>"
}

IMPORTANTE: Si la descripción no tiene suficiente información, genera al menos 3 puntos clave basándote en la información proporcionada.  
No devuelvas ningún texto fuera del JSON.`;

const openai = new OpenAI({
  // baseURL: "https://api.deepseek.com",
  baseURL: "http://localhost:1234/v1/",
  apiKey: "sk-cb409945a45d495d9310f7ccba0b33f9",
});


const results = [];

async function processExcelData() {
  for (const row of data) {
    const userPrompt = JSON.stringify(row, null, 2);
    console.log(userPrompt);
    try {
      const completion = await openai.chat.completions.create({
        messages: [
          { role: "system", content: systemPrompt },
          { role: "user", content: userPrompt },
        ],
        model: "deepseek-chat",
      });

      const response = completion.choices[0].message.content;
      console.log(response);
      const jsonResponse = JSON.parse(response); 
      
 
      results.push({
        ...row,
        Bullet: jsonResponse.bullet.join("\n"), 
        Code: jsonResponse.code,
      });

    } catch (error) {
      console.error("Error procesando la fila:", error);
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