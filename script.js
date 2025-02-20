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

const systemPrompt = `Eres un asistente que solo responde la data generada y nada más. 
Genera un JSON con esta estructura:
{
  "bullet": ["punto 1", "punto 2", "punto 3"],
  "code": "<body>...</body>"
}
La descripción debe convertirse en una lista de puntos clave (bullet) y generar un código HTML en "code".
NO agregues explicaciones ni razonamiento, solo responde con el JSON estructurado.`;


const openai = new OpenAI({
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