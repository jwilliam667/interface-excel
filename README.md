# interface-excel

<!DOCTYPE html>
<html lang="fr">
<head>
  <meta charset="UTF-8">
  <title>Interface Calcul Excel</title>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
  <style>
    body {
      font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
      margin: 0;
      padding: 0;
      background-color: #f9f9f9;
      color: #333;
    }
    header {
      background-color: #0078D7;
      color: white;
      padding: 20px;
      text-align: center;
      box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }
    main {
      padding: 40px;
      max-width: 900px;
      margin: auto;
    }
    h1 {
      margin-top: 0;
    }
    input[type="file"] {
      padding: 10px;
      font-size: 16px;
    }
    button {
      margin: 20px 10px 0 0;
      padding: 10px 20px;
      font-size: 16px;
      cursor: pointer;
      background-color: #0078D7;
      color: white;
      border: none;
      border-radius: 4px;
      transition: background-color 0.3s;
    }
    button:hover {
      background-color: #005fa3;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      margin-top: 30px;
      background-color: white;
      box-shadow: 0 2px 5px rgba(0,0,0,0.05);
    }
    th, td {
      border: 1px solid #ddd;
      padding: 10px;
      text-align: left;
    }
    th {
      background-color: #f2f2f2;
    }
  </style>
</head>
<body>

  <header>
    <h1>Interface Web - Calcul Excel</h1>
  </header>

  <main>
    <input type="file" id="fileInput" accept=".xlsx" />
    <button id="calcButton" disabled>Calculer la somme de la colonne B</button>
    <button id="exportButton" disabled>Exporter les données modifiées</button>
    <div id="output"></div>
  </main>

  <script>
    let jsonData = [];

    document.getElementById('fileInput').addEventListener('change', function(e) {
      const file = e.target.files[0];
      const reader = new FileReader();

      reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

        displayTable(jsonData);
        document.getElementById('calcButton').disabled = false;
        document.getElementById('exportButton').disabled = false;
      };

      reader.readAsArrayBuffer(file);
    });

    document.getElementById('calcButton').addEventListener('click', function() {
      if (jsonData.length > 1) {
        let sum = 0;
        for (let i = 1; i < jsonData.length; i++) {
          const value = parseFloat(jsonData[i][1]); // Colonne B = index 1
          if (!isNaN(value)) sum += value;
        }
        alert("Somme de la colonne B : " + sum);
      }
    });

    document.getElementById('exportButton').addEventListener('click', function() {
      const worksheet = XLSX.utils.aoa_to_sheet(jsonData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Feuille1");

      XLSX.writeFile(workbook, "export_excel.xlsx");
    });

    function displayTable(data) {
      let html = "<table>";
      data.forEach((row, i) => {
        html += "<tr>";
        row.forEach(cell => {
          html += i === 0
            ? `<th>${cell || ""}</th>`
            : `<td>${cell || ""}</td>`;
        });
        html += "</tr>";
      });
      html += "</table>";
      document.getElementById("output").innerHTML = html;
    }
  </script>

</body>
</html>
