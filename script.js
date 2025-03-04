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

const systemPrompt = `
Eres un asistente que genera descripciones de productos en formato JSON. **No devuelvas texto adicional**, solo la respuesta en formato JSON con la estructura exacta indicada abajo.

**Reglas estrictas**:
- La respuesta debe ser **exclusivamente** un JSON válido, sin explicaciones ni otro texto.
- La estructura del JSON **debe ser exactamente esta**:
  {
    "ProductId": "Sin cambios",
    "Name": "Sin cambios",
    "ProductDescription": "Mejorada basada en la descripción proporcionada",
    "bullet": ["Punto clave 1", "Punto clave 2", "Punto clave 3"],
    "code": "<body><h4><strong>Nombre del producto</strong></h4><br/><p><span>Descripción</span></p><br/><br><ul><li><strong>Punto clave 1:</strong> Explicación </li><li><strong>Punto clave 2:</strong> Explicación</li><li><strong>Punto clave 3:</strong> Explicación</li></ul></body>"
  }
- La "ProductDescription" debe mejorar la descripción original y convertir HTML en texto plano.
- El array "bullet" debe contener al menos **3 puntos clave** sobre el producto.
- "code" debe contener una lista <ul> en HTML siguiendo la estructura exacta dada.
- **NO devuelvas texto adicional fuera del JSON**. Si devuelves otro texto, la respuesta será inválida.

**Ejemplo de salida esperada**:
{
  "ProductId": 1234,
  "Name": "Transportador de Mascotas X",
  "ProductDescription": "Transportador seguro y cómodo aprobado por IATA. Funciona como transportador, cama y jaula de entrenamiento.",
  "bullet": [
    "Aprobado por IATA para viajes en avión.",
    "Se convierte en cama o jaula de entrenamiento.",
    "Sujeciones seguras para estabilidad durante el viaje."
  ],
  "code": "<body><h4><strong>Transportador de Mascotas X</strong></h4><br/><p><span>Descripción</span></p><br/><br><ul><li><strong>Aprobado por IATA:</strong> Viajes seguros en avión.</li><li><strong>Versátil:</strong> Transportador, cama o jaula de entrenamiento.</li><li><strong>Diseño seguro:</strong> Sujeciones para estabilidad.</li></ul></body>"
}

**IMPORTANTE**: La salida debe comenzar con { y terminar con }. Si generas algo diferente, la respuesta será rechazada.
`;


const openai = new OpenAI({
  baseURL: "http://localhost:1234/v1/",
  apiKey: "sk-cb409945a45d495d9310f7ccba0b33f9",
});

const results = [];
const maxRetries = 5;

async function retryOperation(row, userPrompt, retries = 0) {
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
      show_reasoning:false,
    });

    const response = completion.choices[0].message.content;

    if (!response || !response.startsWith("{")) {
      console.error(`⚠️ Respuesta no válida en fila:`, row);
      if (retries < maxRetries) {
        console.log(`⏳ Intentando nuevamente en fila... Reintentos restantes: ${maxRetries - retries}`);
        return retryOperation(row, userPrompt, retries + 1); 
      } else {
        console.error(`❌ No se pudo procesar la fila después de ${maxRetries} intentos.`);
        return null; 
      }
    }

    const jsonResponse = JSON.parse(response);


    if (!jsonResponse.bullet || jsonResponse.bullet.length < 3) {
      console.error(`⚠️ "bullet" no válido o insuficiente en fila ${row.ProductId}. Volviendo a solicitar...`);
      if (retries < maxRetries) {
        return retryOperation(row, userPrompt, retries + 1); 
      } else {
        console.error(`❌ No se pudo obtener 'bullet' válido después de ${maxRetries} intentos.`);
        return null;
      }
    }

    return jsonResponse;
  } catch (error) {
    console.error(`⚠️ Error en fila durante la solicitud de OpenAI:`, error);
    if (retries < maxRetries) {
      console.log(`⏳ Intentando nuevamente en fila... Reintentos restantes: ${maxRetries - retries}`);
      return retryOperation(row, userPrompt, retries + 1);
    } else {
      console.error(`❌ No se pudo procesar la fila después de ${maxRetries} intentos.`);
      return null; 
    }
  }
}

async function processExcelData() {
  for (const [index, row] of data.entries()) {
    const userPrompt = JSON.stringify(row, null, 2);
    const jsonResponse = await retryOperation(row, userPrompt); 

    if (jsonResponse) {
      results.push({
        ProductId: row.ProductId,
        Name: String(jsonResponse["Name"]),
        ProductDescription: String(jsonResponse["ProductDescription"]),
        Bullet: jsonResponse["bullet"].join("\n"),
        Code: jsonResponse["code"],
      });

      console.log(`✅ Procesado: ${index + 1} / ${data.length} (${((index + 1) / data.length * 100).toFixed(2)}%)`);
    } else {
      console.error(`❌ Fila ${index + 1} no procesada correctamente después de reintentos.`);
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
