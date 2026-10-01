window.InvestStore = {
  key: "alfa-invest-prototype",
  initial: {
    onboarded: false,
    cash: 2500,
    favorites: ["minfin", "belaruskali"],
    holdings: { minfin: 2, belaruskali: 1, amkodor: 1 },
    orders: [{ id: "AI-1042", asset: "minfin", side: "Покупка", quantity: 2, total: 1192.57, status: "Исполнена" }],
    homeBlocks: { chart: true, favorites: true }
  },
  read() {
    try { return { ...this.initial, ...JSON.parse(localStorage.getItem(this.key)) }; }
    catch { return structuredClone(this.initial); }
  },
  write(state) { localStorage.setItem(this.key, JSON.stringify(state)); },
  reset() { localStorage.removeItem(this.key); },
  toggleFavorite(id) {
    const state = this.read();
    state.favorites = state.favorites.includes(id) ? state.favorites.filter(item => item !== id) : [...state.favorites, id];
    this.write(state);
    return state;
  },
  placeOrder({ asset, side, quantity }) {
    const state = this.read();
    const item = InvestData.assets.find(entry => entry.id === asset);
    const subtotal = item.price * quantity;
    const commission = subtotal * .003;
    const total = subtotal + commission;
    if (side === "Покупка") {
      state.cash -= total;
      state.holdings[asset] = (state.holdings[asset] || 0) + quantity;
    } else {
      state.cash += subtotal - commission;
      state.holdings[asset] = Math.max(0, (state.holdings[asset] || 0) - quantity);
    }
    const order = { id: `AI-${Date.now().toString().slice(-5)}`, asset, side, quantity, total, status: "Исполнена" };
    state.orders = [order, ...state.orders];
    this.write(state);
    return { order, subtotal, commission, total };
  },
  cancelOrder(id) {
    const state = this.read();
    state.orders = state.orders.map(order => order.id === id ? { ...order, status: "Отменена" } : order);
    this.write(state);
  }
};
