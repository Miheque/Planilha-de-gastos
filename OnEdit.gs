function onEdit(e) {
  if (!e || !e.range) return;

  const sheetName = e.range.getSheet().getName();

  const abasDividas = ["MARIA", "PRISCILA", "ISAEL", "RAYLAN","REYNAN", "EU", "SHOPEE", "ISAEL SHOPEE", "CARTÃO M-PAGO", "CRÉDITO M-PAGO", "CARTÃO PAN"];
  const abasSaldos   = ["SALDO", "JOSÉ'S MONEY", "REYNAN'S MONEY"];
  
  if (abasDividas.includes(sheetName)) {
    handleDividas(e);
    return;
  }

  if (abasSaldos.includes(sheetName)) {
    handleSaldos(e);
    return;
  }
}
