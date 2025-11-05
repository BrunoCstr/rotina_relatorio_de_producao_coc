import dotenv from "dotenv";
dotenv.config();

/**
 * Autentica no SGCOR e retorna o token
 * @returns {Promise<{data: {token: string}}>}
 */
export async function getAuth() {
  const url = "https://apirest.gruposgcor.com.br/api/login";

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      email: process.env.SGCOR_USERNAME,
      senha: process.env.SGCOR_PASSWORD,
    }),
  });

  if (!response.ok) {
    throw new Error(`Erro na autenticação: ${response.status}`);
  }

  const data = await response.json();
  return data;
}

/**
 * URL base da API do SGCOR
 */
export const infoAPI = {
  url: process.env.SGCOR_API_URL,
};
