// URLクエリ文字列からパラメータを取得する関数
function getQueryParams() {
    const params = new URLSearchParams(window.location.search);
    const markers = params.get('markers');
    return markers ? JSON.parse(markers) : [];
}

// マーカーを表示する関数
function renderMarkers(markers) {
    const container = document.getElementById('image-container');
    markers.forEach((markerData, index) => {
        const marker = createMarker(markerData.x * container.clientWidth, markerData.y * container.clientHeight, index);
        container.appendChild(marker);
    });
}

// 新しいマーカーを作成する関数
function createMarker(x, y, index) {
    const marker = document.createElement('div');
    marker.className = 'marker';
    marker.style.left = `${x}px`;
    marker.style.top = `${y}px`;
    marker.dataset.index = index;

    // マーカークリックイベントリスナーを追加して削除
    marker.addEventListener('click', function(event) {
        event.stopPropagation(); // 親要素へのクリックイベント伝播を防ぐ
        const index = parseInt(marker.dataset.index, 10);
        const markers = getQueryParams();
        markers.splice(index, 1);
        saveMarkersToQueryParams(markers);
        marker.remove();

        // インデックスを更新
        updateMarkerIndices();
    });

    return marker;
}

// クエリパラメータにマーカーを保存する関数
function saveMarkersToQueryParams(markers) {
    const params = new URLSearchParams(window.location.search);
    params.set('markers', JSON.stringify(markers));
    history.replaceState(null, '', `${window.location.pathname}?${params.toString()}`);
}

// インデックスを更新する関数
function updateMarkerIndices() {
    const markers = document.querySelectorAll('.marker');
    markers.forEach((marker, index) => {
        marker.dataset.index = index;
    });
}

document.getElementById('image-container').addEventListener('click', function(event) {
    // クリックされた要素がマーカーでないことを確認
    if (event.target.classList.contains('marker')) return;

    const container = event.currentTarget;
    const rect = container.getBoundingClientRect();
    const x = (event.clientX - rect.left) / rect.width;
    const y = (event.clientY - rect.top) / rect.height;

    const marker = createMarker(x * container.clientWidth, y * container.clientHeight, getQueryParams().length);
    container.appendChild(marker);

    const markers = getQueryParams();
    markers.push({ x, y });
    saveMarkersToQueryParams(markers);
});

// ページ読み込み時にクエリパラメータからマーカーを取得して表示
window.addEventListener('load', function() {
    const markers = getQueryParams();
    renderMarkers(markers);
});
