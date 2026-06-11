import type { ParsedAssembly } from './parser';

export async function analyzeBOMImage(
  base64DataUrl: string,
  apiKey: string
): Promise<ParsedAssembly[]> {
  // Extract clean base64 data and mime type
  const match = base64DataUrl.match(/^data:(image\/[a-zA-Z+]+);base64,(.+)$/);
  if (!match) {
    throw new Error('სურათის მონაცემების ფორმატი არასწორია. აუცილებელია base64 გამოსახულება.');
  }

  const mimeType = match[1];
  const base64Data = match[2];

  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${apiKey}`;

  const prompt = `
You are an expert engineer and shop draftsman. Your task is to extract a Bill of Materials (BOM) or Weldment Cut List table from the uploaded screenshot.

Analyze the image carefully. Identify all assemblies (constructions) and the cut items belonging to each.

Important rules:
1. Identify parent assemblies (constructions):
   - They usually have integer positions (e.g., 1, 2) or names like "მოაჯირი..." (Handrail) and a quantity (e.g., 1, 31). They do NOT have a profile/cut size in the description.
   - Extract their Name (დასახელება) and Qty (რაოდ).
2. Identify child items (cut members):
   - They have decimal positions (e.g., 1.1, 1.2, 2.1, 2.2).
   - Extract their Position (პოზიცია), Qty (რაოდ, e.g. 15), Profile Description (აღწერა, e.g., "50 x 50 x 3.2", "20 x 20 x 2.0"), and Single Cut Length (სიგრძე, e.g., 1090).
   - Convert length to millimeters (numbers only). If it's in meters (e.g., 1.09), convert it to mm (e.g., 1090).
   - DO NOT confuse the single cut length (სიგრძე) with the total length (ჯამ. სიგრძე). We need the SINGLE CUT length.
3. If the image has column headers in Georgian (like the provided one):
   - პოზიცია = Position
   - რაოდ = Quantity (Qty)
   - დასახელება = Name
   - აღწერა = Description / Profile
   - სიგრძე = Single Length
   - ჯამ. სიგრძე = Total Length (ignore this, we will compute our own total, or only use it to verify)
4. Also look out for sheet metal/plate materials (e.g., "ფურცელი", "plate", "sheet", "S=3", "thickness 2mm"). Put them in the list as well; their description should indicate it is a sheet/plate.
5. Return the result strictly as a JSON object matching this TypeScript interface:

interface Response {
  constructions: {
    name: string;
    position: string;
    qty: number;
    items: {
      position: string;
      qty: number;
      description: string; // e.g. "50 x 50 x 3.2"
      length: number; // in mm
    }[];
  }[];
}
`;

  const requestBody = {
    contents: [
      {
        parts: [
          { text: prompt },
          {
            inlineData: {
              mimeType: mimeType,
              data: base64Data,
            },
          },
        ],
      },
    ],
    generationConfig: {
      responseMimeType: 'application/json',
      temperature: 0.1,
    },
  };

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(requestBody),
  });

  if (!response.ok) {
    const errorData = await response.json().catch(() => ({}));
    const errorMessage = errorData?.error?.message || response.statusText;
    throw new Error(`Gemini API შეცდომა: ${errorMessage}`);
  }

  const result = await response.json();
  const textContent = result.candidates?.[0]?.content?.parts?.[0]?.text;
  
  if (!textContent) {
    throw new Error('Gemini მოდელიდან პასუხი არ დაბრუნებულა.');
  }

  try {
    const parsed = JSON.parse(textContent);
    if (!parsed.constructions || !Array.isArray(parsed.constructions)) {
      throw new Error('პასუხის ფორმატში აკლია "constructions" მასივი.');
    }
    return parsed.constructions;
  } catch (err) {
    console.error('Failed to parse Gemini output:', textContent, err);
    throw new Error('AI პასუხის სწორ BOM JSON-ად გაანალიზება ვერ მოხერხდა.');
  }
}
