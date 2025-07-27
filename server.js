// Importar las librerías necesarias
const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const path = require('path');

const app = express();
// Render te dará un puerto en la variable de entorno PORT
const PORT = process.env.PORT || 3000;

// Configuración para servir los archivos estáticos (html, css, js) desde la carpeta 'public'
app.use(express.static(path.join(__dirname, 'public')));

// Configuración de Multer para manejar la subida de archivos en memoria
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Definición de las palabras a ignorar (stopwords)
const STOPWORDS = [
    'de', 'la', 'que', 'el', 'en', 'y', 'a', 'los', 'del', 'se', 'las', 'por', 'un',
    'para', 'con', 'no', 'una', 'su', 'al', 'lo', 'como', 'más', 'pero', 'sus',
    'le', 'ya', 'o', 'este', 'ha', 'me', 'si', 'sin', 'sobre', 'este', 'muy',
    'cuando', 'también', 'hasta', 'hay', 'donde', 'quien', 'desde', 'todo',
    'nos', 'durante', 'uno', 'ni', 'contra', 'ese', 'eso', 'mi', 'qué', 'e', 'son',
    'fue', 'muy', 'gracias', 'hola', 'buen', 'dia'
];

// Función para limpiar y obtener palabras de un texto
function getWordsFromString(text) {
    if (!text || typeof text !== 'string') return [];
    const textLower = text.toLowerCase();
    const words = textLower.match(/\b(\w+)\b/g) || []; // Obtener solo palabras
    return words.filter(word => !STOPWORDS.includes(word) && word.length > 2);
}

// La ruta principal de nuestra API que procesará el archivo
app.post('/procesar', upload.single('archivoExcel'), (req, res) => {
    try {
        if (!req.file) {
            return res.status(400).json({ success: false, message: 'No se subió ningún archivo.' });
        }

        // Leer el archivo Excel desde el buffer en memoria
        const workbook = xlsx.read(req.file.buffer, { type: 'buffer' });
        const sheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[sheetName];
        const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

        // Misma estructura de datos que en la versión PHP
        let processedData = {
            general: { muy_positivas: 0, positivas: 0, negativas: 0, muy_negativas: 0, total: 0 },
            porDia: {},
            porHora: Array.from({ length: 24 }, () => ({ muy_positivas: 0, positivas: 0, negativas: 0, muy_negativas: 0, total: 0 })),
            porSector: {},
            comentarios: { positivos: [], negativos: [] },
            fechas: []
        };
        
        let uniqueDates = {};
        const DIAS_SEMANA = ['Domingo', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado'];

        // Empezar desde la fila 1 para saltar el encabezado
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            if (!row[0]) continue;

            // Excel maneja las fechas como números, necesitamos convertirlos
            const excelDate = parseFloat(row[0]);
            const jsDate = new Date((excelDate - (25567 + 2)) * 86400 * 1000); // Corrección para zona horaria y formato Excel
            
            if (isNaN(jsDate.getTime())) continue; // Saltar si la fecha no es válida
            
            const diaSemana = DIAS_SEMANA[jsDate.getDay()];
            const fecha = jsDate.toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' });
            const hora = jsDate.getHours();
            
            uniqueDates[fecha] = true;

            const sector = (row[2] || '').trim();
            const ubicacion = (row[3] || '').trim();
            const sectorKey = `${sector} - ${ubicacion}`;
            const calificacionDesc = (row[7] || '').trim();
            const comentario = (row[4] || '');

            // Inicializar contadores si no existen
            if (!processedData.porDia[diaSemana]) {
                processedData.porDia[diaSemana] = { muy_positivas: 0, positivas: 0, negativas: 0, muy_negativas: 0, total: 0 };
            }
            if (!processedData.porSector[sectorKey]) {
                processedData.porSector[sectorKey] = { muy_positivas: 0, positivas: 0, negativas: 0, muy_negativas: 0, total: 0, comentarios: [] };
            }

            processedData.general.total++;
            processedData.porDia[diaSemana].total++;
            processedData.porHora[hora].total++;
            processedData.porSector[sectorKey].total++;

            // Contar calificaciones
            switch (calificacionDesc) {
                case 'Muy Positiva':
                    processedData.general.muy_positivas++;
                    processedData.porDia[diaSemana].muy_positivas++;
                    processedData.porHora[hora].muy_positivas++;
                    processedData.porSector[sectorKey].muy_positivas++;
                    processedData.comentarios.positivos.push(...getWordsFromString(comentario));
                    break;
                case 'Positiva':
                    processedData.general.positivas++;
                    processedData.porDia[diaSemana].positivas++;
                    processedData.porHora[hora].positivas++;
                    processedData.porSector[sectorKey].positivas++;
                    processedData.comentarios.positivos.push(...getWordsFromString(comentario));
                    break;
                case 'Negativa':
                    processedData.general.negativas++;
                    processedData.porDia[diaSemana].negativas++;
                    processedData.porHora[hora].negativas++;
                    processedData.porSector[sectorKey].negativas++;
                    processedData.comentarios.negativos.push(...getWordsFromString(comentario));
                    break;
                case 'Muy Negativa':
                    processedData.general.muy_negativas++;
                    processedData.porDia[diaSemana].muy_negativas++;
                    processedData.porHora[hora].muy_negativas++;
                    processedData.porSector[sectorKey].muy_negativas++;
                    processedData.comentarios.negativos.push(...getWordsFromString(comentario));
                    break;
            }
        }
        
        processedData.fechas = Object.keys(uniqueDates);

        const calculateNPS = (stats) => {
            if (stats.total === 0) return 0;
            const promotores = stats.muy_positivas; // Solo las muy positivas son promotores
            const detractores = stats.negativas + stats.muy_negativas;
            return Math.round(((promotores - detractores) / stats.total) * 100 * 100) / 100; // NPS con 2 decimales
        };

        processedData.general.nps = calculateNPS(processedData.general);
        for (const dia in processedData.porDia) {
            processedData.porDia[dia].nps = calculateNPS(processedData.porDia[dia]);
        }
        for (const hora in processedData.porHora) {
            processedData.porHora[hora].nps = calculateNPS(processedData.porHora[hora]);
        }
        for (const sector in processedData.porSector) {
            processedData.porSector[sector].nps = calculateNPS(processedData.porSector[sector]);
        }

        // Preparar nubes de palabras
        const countWords = (arr) => arr.reduce((acc, word) => { acc[word] = (acc[word] || 0) + 1; return acc; }, {});
        processedData.nubes = {
            positiva: countWords(processedData.comentarios.positivos),
            negativa: countWords(processedData.comentarios.negativos)
        };
        
        // Devolver el resultado
        res.json({ success: true, data: processedData });

    } catch (error) {
        console.error('Error procesando el archivo:', error);
        res.status(500).json({ success: false, message: 'Error interno del servidor.' });
    }
});

// Iniciar el servidor
app.listen(PORT, () => {
    console.log(`Servidor corriendo en el puerto ${PORT}`);
});
