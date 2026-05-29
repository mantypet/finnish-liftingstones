const map = L.map('map').setView([64.5, 26.0], 5);

L.tileLayer('https://{s}.tile.opentopomap.org/{z}/{x}/{y}.png', {
  attribution: '© <a href="https://opentopomap.org">OpenTopoMap</a> (<a href="https://creativecommons.org/licenses/by-sa/3.0/">CC-BY-SA</a>)',
  maxZoom: 17
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
        <tr><td>Location</td><td>${stone.location}</td></tr>
        <tr><td>Weight</td><td>${stone.weight_kg} kg</td></tr>
        <tr><td>Historical Stone</td><td>${stone.historical}</td></tr>
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
      `${data.stones.length} stones`;
  })
  .catch(err => console.error('Failed to load stones.json:', err));
