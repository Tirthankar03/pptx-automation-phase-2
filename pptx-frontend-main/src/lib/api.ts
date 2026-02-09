const BASE_URL = "/api";

export interface ProjectUpdatePayload {
  type: "project_update";
  title: string;
  columns: string[];
  content: string[][];
}

export const generatePptxFromJson = async (payload: ProjectUpdatePayload): Promise<Blob> => {
  const response = await fetch(`${BASE_URL}/generate-pptx`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "ngrok-skip-browser-warning": "true",
    },
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    throw new Error(`Failed to generate PPTX: ${response.statusText}`);
  }

  return response.blob();
};

export const generatePptxFromExcel = async (file: File): Promise<Blob> => {
  const formData = new FormData();
  formData.append("file", file);

  const response = await fetch(`${BASE_URL}/generate-pptx-from-excel`, {
    method: "POST",
    headers: {
      "ngrok-skip-browser-warning": "true",
    },
    body: formData,
  });

  if (!response.ok) {
    throw new Error(`Failed to generate PPTX from Excel: ${response.statusText}`);
  }

  return response.blob();
};

export const downloadTemplate = async (): Promise<Blob> => {
  const response = await fetch(`${BASE_URL}/download-template`, {
    method: "GET",
    headers: {
      "ngrok-skip-browser-warning": "true",
    },
  });

  if (!response.ok) {
    throw new Error(`Failed to download template: ${response.statusText}`);
  }

  return response.blob();
};

export const downloadBlob = (blob: Blob, filename: string) => {
  const url = window.URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  window.URL.revokeObjectURL(url);
};
