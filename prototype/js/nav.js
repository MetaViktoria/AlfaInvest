window.InvestNav = {
  items: [
    ["home.html", "⌂", "Главная", "home"],
    ["portfolio.html", "◔", "Портфель", "portfolio"],
    ["trade.html", "▤", "Каталог", "trade"],
    ["help.html", "?", "Помощь", "help"],
    ["profile.html", "◉", "Профиль", "profile"]
  ],
  render(active) {
    const links = this.items.map(([href, icon, label, id]) =>
      `<a href="${href}" class="${active === id ? "active" : ""}"><span class="nav-icon">${icon}</span><span>${label}</span></a>`
    ).join("");
    document.body.insertAdjacentHTML("afterbegin", `
      <div class="app-shell">
        <aside class="side">
          <a class="brand" href="home.html"><span class="brand-mark">A</span>Alfa Invest</a>
          <nav class="nav">${links}</nav>
          <div class="side-bottom">Тестовое задание. Интерактивный прототип.<br>Не является инвестиционной рекомендацией</div>
        </aside>
        <main class="main" id="page"></main>
      </div>
      <nav class="bottom-nav">${links}</nav>
    `);
  },
  assetRows(ids, withAmount = false) {
    const state = InvestStore.read();
    return ids.map(id => {
      const item = InvestData.assetById(id);
      if (!item) return "";
      const quantity = state.holdings[id] || 0;
      return `<a class="asset" href="instrument.html?id=${item.id}">
        <span class="ticker">${item.ticker}</span>
        <span class="asset-name"><b>${item.name}</b><span>${item.type} · ${item.sector}${withAmount ? ` · ${quantity} шт.` : ""}</span></span>
        <svg class="spark" viewBox="0 0 68 24" preserveAspectRatio="none"><polyline points="0,18 9,16 18,19 28,12 38,14 48,8 58,10 68,4"/></svg>
        <span><b>${InvestData.format(item.price)}</b><br><small class="${item.change >= 0 ? "positive" : "negative"}">${item.change >= 0 ? "+" : ""}${item.change}%</small></span>
      </a>`;
    }).join("");
  }
};
