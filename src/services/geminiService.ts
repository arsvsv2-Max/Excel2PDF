import { GoogleGenAI, Type } from "@google/genai";

export interface MappingResult {
  pdfField: string;
  excelHeader: string;
}

export const getSmartMapping = async (
  pdfFields: string[],
  excelHeaders: string[]
): Promise<MappingResult[]> => {
  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    throw new Error("GEMINI_API_KEY is not set");
  }

  const ai = new GoogleGenAI({ apiKey });
  
  const prompt = `
    I have a PDF form with these fields: [${pdfFields.join(", ")}]
    And an Excel spreadsheet with these headers: [${excelHeaders.join(", ")}]
    
    Please analyze the semantic similarity and return a mapping of which Excel header should fill which PDF field.
    Return only a JSON array of objects with "pdfField" and "excelHeader" keys.
    If no good match is found for a field, omit it from the array.
  `;

  const response = await ai.models.generateContent({
    model: "gemini-3-flash-preview",
    contents: prompt,
    config: {
      responseMimeType: "application/json",
      responseSchema: {
        type: Type.ARRAY,
        items: {
          type: Type.OBJECT,
          properties: {
            pdfField: { type: Type.STRING },
            excelHeader: { type: Type.STRING },
          },
          required: ["pdfField", "excelHeader"],
        },
      },
    },
  });

  try {
    return JSON.parse(response.text || "[]");
  } catch (e) {
    console.error("Failed to parse AI response", e);
    return [];
  }
};
