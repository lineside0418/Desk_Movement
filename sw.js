// インストール要件を満たすための最低限のService Worker
self.addEventListener('install', (e) => {
  self.skipWaiting();
});

self.addEventListener('fetch', (e) => {
  // 常にネットワークから最新のデータを取得するため、ここでは何もしない
});