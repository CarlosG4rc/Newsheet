function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Documento')
      .addItem('NewSheet', 'newSheet')

      .addToUi();
}
function newSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var lastrowbd = ss.getDataRange().getNumRows();
  var lastcolbd = ss.getDataRange().getNumColumns();
  var column = ss.getDataRange();
  var value = column.getValues();
  var b = 3;
  for(var i = 1; i < lastrowbd; i++)//Recorremos todas las filas 
  {
    var a = i-1;
    var nom = value[a] && value [a][b]
    if(value[i][b] != nom)//Comparamos nombre introducido con el dato anterior
    {
       var nombre = value[i] && value [i][b];
       var newsheet = ss.getSheetByName("" + nombre);
       newsheet = ss.insertSheet();
       newsheet.setName("" + nombre);
       //newsheet.getRange("A" + i + ":Z" + i).copyTo(newsheet.getRange("A" + i + ":Z" + i));
       newsheet.getRange("A5").setValue('=FILTER( Respuestas3!A2:Z1000 ; Respuestas3!D2:D1000="'+ nombre +'")');
       
       newsheet.getRange("E1").setValue('Muy bueno: ');
       newsheet.getRange("F1").setFormula("=COUNTIF(F5:F2000;\"Muy bueno\")");
       newsheet.getRange("E2").setValue('Bueno: ');
       newsheet.getRange("F2").setFormula("=COUNTIF(F5:F2000;\"Bueno\")");
       newsheet.getRange("E3").setValue('Regular: ');
       newsheet.getRange("F3").setFormula("=COUNTIF(F5:F2000;\"Regular\")");
       newsheet.getRange("E4").setValue('Necesita mejorar: ');
       newsheet.getRange("F4").setFormula("=COUNTIF(F5:F2000;\"Necesita mejorar\")");
       
       newsheet.getRange("G1").setFormula("=COUNTIF(G5:G2000;\"Muy bueno\")");
       newsheet.getRange("G2").setFormula("=COUNTIF(G5:G2000;\"Bueno\")");
       newsheet.getRange("G3").setFormula("=COUNTIF(G5:G2000;\"Regular\")");
       newsheet.getRange("G4").setFormula("=COUNTIF(G5:G2000;\"Necesita mejorar\")");
       
       newsheet.getRange("H1").setFormula("=COUNTIF(H5:H2000;\"Muy bueno\")");
       newsheet.getRange("H2").setFormula("=COUNTIF(H5:H2000;\"Bueno\")");
       newsheet.getRange("H3").setFormula("=COUNTIF(H5:H2000;\"Regular\")");
       newsheet.getRange("H4").setFormula("=COUNTIF(H5:H2000;\"Necesita mejorar\")");
       
       newsheet.getRange("I1").setFormula("=COUNTIF(I5:I2000;\"Muy bueno\")");
       newsheet.getRange("I2").setFormula("=COUNTIF(I5:I2000;\"Bueno\")");
       newsheet.getRange("I3").setFormula("=COUNTIF(I5:I2000;\"Regular\")");
       newsheet.getRange("I4").setFormula("=COUNTIF(I5:I2000;\"Necesita mejorar\")");
       
       newsheet.getRange("J1").setFormula("=COUNTIF(J5:J2000;\"Muy bueno\")");
       newsheet.getRange("J2").setFormula("=COUNTIF(J5:J2000;\"Bueno\")");
       newsheet.getRange("J3").setFormula("=COUNTIF(J5:J2000;\"Regular\")");
       newsheet.getRange("J4").setFormula("=COUNTIF(J5:J2000;\"Necesita mejorar\")");
       
       newsheet.getRange("K1").setFormula("=COUNTIF(K5:K2000;\"Muy bueno\")");
       newsheet.getRange("K2").setFormula("=COUNTIF(K5:K2000;\"Bueno\")");
       newsheet.getRange("K3").setFormula("=COUNTIF(K5:K2000;\"Regular\")");
       newsheet.getRange("K4").setFormula("=COUNTIF(K5:K2000;\"Necesita mejorar\")");
       
       newsheet.getRange("L1").setFormula("=COUNTIF(L5:L2000;\"Muy bueno\")");
       newsheet.getRange("L2").setFormula("=COUNTIF(L5:L2000;\"Bueno\")");
       newsheet.getRange("L3").setFormula("=COUNTIF(L5:L2000;\"Regular\")");
       newsheet.getRange("L4").setFormula("=COUNTIF(L5:L2000;\"Necesita mejorar\")");
       
       newsheet.getRange("M1").setFormula("=COUNTIF(M5:M2000;\"Muy bueno\")");
       newsheet.getRange("M2").setFormula("=COUNTIF(M5:M2000;\"Bueno\")");
       newsheet.getRange("M3").setFormula("=COUNTIF(M5:M2000;\"Regular\")");
       newsheet.getRange("M4").setFormula("=COUNTIF(M5:M2000;\"Necesita mejorar\")");
       
       newsheet.getRange("N1").setFormula("=COUNTIF(N5:N2000;\"Muy bueno\")");
       newsheet.getRange("N2").setFormula("=COUNTIF(N5:N2000;\"Bueno\")");
       newsheet.getRange("N3").setFormula("=COUNTIF(N5:N2000;\"Regular\")");
       newsheet.getRange("N4").setFormula("=COUNTIF(N5:N2000;\"Necesita mejorar\")");
       
       newsheet.getRange("O1").setFormula("=COUNTIF(O5:O2000;\"Muy bueno\")");
       newsheet.getRange("O2").setFormula("=COUNTIF(O5:O2000;\"Bueno\")");
       newsheet.getRange("O3").setFormula("=COUNTIF(O5:O2000;\"Regular\")");
       newsheet.getRange("O4").setFormula("=COUNTIF(O5:O2000;\"Necesita mejorar\")");
       
       newsheet.getRange("P1").setFormula("=COUNTIF(P5:P2000;\"Muy bueno\")");
       newsheet.getRange("P2").setFormula("=COUNTIF(P5:P2000;\"Bueno\")");
       newsheet.getRange("P3").setFormula("=COUNTIF(P5:P2000;\"Regular\")");
       newsheet.getRange("P4").setFormula("=COUNTIF(P5:P2000;\"Necesita mejorar\")");
       
       newsheet.getRange("Q1").setFormula("=COUNTIF(Q5:Q2000;\"Muy bueno\")");
       newsheet.getRange("Q2").setFormula("=COUNTIF(Q5:Q2000;\"Bueno\")");
       newsheet.getRange("Q3").setFormula("=COUNTIF(Q5:Q2000;\"Regular\")");
       newsheet.getRange("Q4").setFormula("=COUNTIF(Q5:Q2000;\"Necesita mejorar\")");
       
       newsheet.getRange("R1").setFormula("=COUNTIF(R5:R2000;\"Muy bueno\")");
       newsheet.getRange("R2").setFormula("=COUNTIF(R5:R2000;\"Bueno\")");
       newsheet.getRange("R3").setFormula("=COUNTIF(R5:R2000;\"Regular\")");
       newsheet.getRange("R4").setFormula("=COUNTIF(R5:R2000;\"Necesita mejorar\")");
       
       newsheet.getRange("S1").setFormula("=COUNTIF(S5:S2000;\"Muy bueno\")");
       newsheet.getRange("S2").setFormula("=COUNTIF(S5:S2000;\"Bueno\")");
       newsheet.getRange("S3").setFormula("=COUNTIF(S5:S2000;\"Regular\")");
       newsheet.getRange("S4").setFormula("=COUNTIF(S5:S2000;\"Necesita mejorar\")");
       
       newsheet.getRange("T1").setFormula("=COUNTIF(T5:T2000;\"Muy bueno\")");
       newsheet.getRange("T2").setFormula("=COUNTIF(T5:T2000;\"Bueno\")");
       newsheet.getRange("T3").setFormula("=COUNTIF(T5:T2000;\"Regular\")");
       newsheet.getRange("T4").setFormula("=COUNTIF(T5:T2000;\"Necesita mejorar\")");
       
       newsheet.getRange("U1").setFormula("=COUNTIF(U5:U2000;\"Muy bueno\")");
       newsheet.getRange("U2").setFormula("=COUNTIF(U5:U2000;\"Bueno\")");
       newsheet.getRange("U3").setFormula("=COUNTIF(U5:U2000;\"Regular\")");
       newsheet.getRange("U4").setFormula("=COUNTIF(U5:U2000;\"Necesita mejorar\")");
       
       newsheet.getRange("V1").setFormula("=COUNTIF(V5:V2000;\"Sí\")");
       newsheet.getRange("V2").setFormula("=COUNTIF(V5:V2000;\"No\")"); 
       
       newsheet.getRange("X1").setValue('Promedio');
       newsheet.getRange("X2").setFormula("=AVERAGE(X5:X2000)");
    }
  }
}