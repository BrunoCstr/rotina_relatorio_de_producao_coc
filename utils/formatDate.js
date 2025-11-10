/**
 * Formata uma data com base em um offset de dias relativo a hoje
 * @param {number} offset - Número de dias (negativo para datas passadas, positivo para futuras)
 * @returns {string} Data formatada como YYYY-MM-DD
 */
export function formatDate(offset = 0) {
  const date = new Date();
  date.setDate(date.getDate() + offset);
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}






