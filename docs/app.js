const map = L.map('map').setView([64.5, 26.0], 5);

L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
  attribution: '© <a href="https://www.openstreetmap.org/copyright">OpenStreetMap</a> contributors',
  maxZoom: 19
}).addTo(map);

const stoneIcon = L.divIcon({
  className: '',
  html: '<div class="stone-marker">🪨</div>',
  iconSize: [30, 30],
  iconAnchor: [15, 15],
  popupAnchor: [0, -15]
});

function buildPopup(stone) {
  const tm35 = stone.northing && stone.easting
    ? `<tr><td>TM35FIN</td><td>N ${stone.northing.toLocaleString('fi-FI')} E ${stone.easting.toLocaleString('fi-FI')}</td></tr>`
    : '';
  const notes = stone.notes
    ? `<tr><td colspan="2" class="popup-notes">${stone.notes}</td></tr>`
    : '';
  return `
    <div class="popup-content">
      <h3>${stone.name}</h3>
      <table class="popup-table">
        <tr><td>Sijainti</td><td>${stone.location}</td></tr>
        <tr><td>Paino</td><td>${stone.weight_kg} kg</td></tr>
        ${tm35}
        ${notes}
      </table>
    </div>`;
}

fetch('stones.json')
  .then(r => r.json())
  .then(data => {
    data.stones.forEach(stone => {
      L.marker([stone.lat, stone.lon], { icon: stoneIcon })
        .addTo(map)
        .bindPopup(buildPopup(stone), { maxWidth: 280 });
    });

    document.getElementById('count').textContent =
      `${data.stones.length} kiveä`;
  })
  .catch(err => console.error('Failed to load stones.json:', err));
