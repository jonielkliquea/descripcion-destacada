import OpenAI from "openai";
import fs from 'fs';
import path from "path";
import * as XLSX from 'xlsx';

try {
  const filePath = path.resolve('src', 'Booktt.xlsx');
  const fileBuffer = fs.readFileSync('./src/Booktt.xlsx');
  const workbook = XLSX.read(fileBuffer, { type: 'buffer' });

  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = XLSX.utils.sheet_to_json(worksheet);


  const systemPromp = `Eres un asistente que solo responde la data generada y nada mas, la data la vas a general en base a el siguiente formto: 
vas a crear a partir de el json dado un json con los parametros dados + bullet donde colocaras los dateos de descripcion principal y con ellos crearas el siguiente parametro code con la siguiente estructura html:
"<body>
    <h1>Descripción del Producto</h1>
    <ul>
        <li><strong>Silueta semi ajustada:</strong> Realza las curvas.</li>
        <li><strong>Escote cruzado en el frente con canal:</strong> Agrega sofisticación.</li>
        <li><strong>Cargaderas ajustables y cut out delantero:</strong> Para un toque atrevido y elegante.</li>
        <li><strong>Confeccionado con materiales de alta calidad:</strong> Ideal para cualquier ocasión.</li>
    </ul>
</body>"  como puedes ver de el texto dado hace una lista de puntos importantes para posteriormente crear un trozo de html esto proviene de un excel y lo genrearemos en formato que podamos tomar la respuesta y agregarla en un excel con js sin ninguna otra data adicional`;

const openai = new OpenAI({
  baseURL: "http://localhost:1234/v1/", 
  apiKey: "sk-cb409945a45d495d9310f7ccba0b33f9",
});

async function main() {
  const userPrompt = JSON.stringify(data, null, 2);
  console.log(userPrompt);
  
  const completion = await openai.chat.completions.create({
    messages: [
      { role: "system", content: systemPromp },
      { role: "user", content: userPrompt }
    ],
    model: "deepseek-chat",
  });

  console.log("Respuesta de DeepSeek:", completion.choices[0].message.content);
}

main();


} catch (error) {
  console.error("Error al leer el archivo:", error);
}
