const syllables = [
  "pa", "pe", "pi", "po", "pu",
  "la", "le", "li", "lo", "lu",
  "ma", "me", "mi", "mo", "mu",
  "na", "ne", "ni", "no", "nu"
];

const table = document.getElementById('syllableTable');
let intervalId;
const selectedCells = {};

function resetGame() {
  const cells = table.getElementsByTagName('td');
  Array.from(cells).forEach(cell => {
    cell.classList.remove('winner');
    cell.classList.remove('highlight');
    cell.style.border = '';  // Restablecer el color del borde
  });
  Object.keys(selectedCells).forEach(key => delete selectedCells[key]);  // Limpiar el objeto selectedCells
}

function changeFont(fontFamily, fontSize) {
  const cells = table.getElementsByTagName('td');
  Array.from(cells).forEach(cell => {
    cell.style.fontFamily = fontFamily;
    cell.style.fontSize = fontSize;
  });
}

function adjustFontSize() {
  const tableWidth = table.offsetWidth;
  const fontSize = Math.min(tableWidth / 25, 36);  // Ajusta los valores según sea necesario
  changeFont(undefined, fontSize + 'px');
}

// Llama a adjustFontSize cuando la página se carga y se redimensiona
window.addEventListener('load', adjustFontSize);
window.addEventListener('resize', adjustFontSize);

// Llenar la tabla con las sílabas
for (let i = 0; i < 4; i++) {
  const row = table.insertRow();
  for (let j = 0; j < 5; j++) {
    const cell = row.insertCell();
    cell.textContent = syllables[i * 5 + j];
  }
}

function startRandom() {
  const cells = table.getElementsByTagName('td');

  // Verificar si todas las celdas han sido seleccionadas
  if (Object.keys(selectedCells).length === cells.length) {
    resetGame();
  }

  if (intervalId) {
    clearInterval(intervalId);
  }

  // Resetear las clases winner y highlight de todas las celdas
  Array.from(cells).forEach(cell => {
    cell.classList.remove('winner');
    cell.classList.remove('highlight');
  });

  intervalId = setInterval(() => {
    let randomIndex = Math.floor(Math.random() * cells.length);
    // Verificar si la celda ya ha sido seleccionada, si es así, obtener otro índice aleatorio
    while (selectedCells[randomIndex]) {
      randomIndex = Math.floor(Math.random() * cells.length);
    }
    Array.from(cells).forEach(cell => cell.classList.remove('highlight'));
    cells[randomIndex].classList.add('highlight');
  }, 100);

  setTimeout(() => {
    clearInterval(intervalId);
    // Encuentra la celda ganadora y añade la clase winner
    const winnerCell = Array.from(cells).find(cell => cell.classList.contains('highlight'));
    if (winnerCell) {
      winnerCell.classList.add('winner');
      winnerCell.style.border = '5px solid green';  // Cambiar el color del borde de la celda ganadora
      const winnerIndex = Array.from(cells).indexOf(winnerCell);
      selectedCells[winnerIndex] = true;
    }
  }, 5000);
}

function resetGame() {
  const cells = table.getElementsByTagName('td');
  Array.from(cells).forEach(cell => {
    cell.classList.remove('winner');
    cell.classList.remove('highlight');
    cell.style.border = '5px solid #d32f2f';  // Restablecer el color del borde
  });
  Object.keys(selectedCells).forEach(key => delete selectedCells[key]);  // Limpiar el objeto selectedCells
}
