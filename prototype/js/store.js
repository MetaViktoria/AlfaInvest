window.InvestStore = {
  key: "alfa-invest-prototype",
  initial: {
    onboarded: false,
    goal: "Накопить на цель",
    experience: "Я начинаю",
    startAmount: "50 BYN",
    cash: 2500,
    favorites: ["minfin", "belaruskali"],
    holdings: { minfin: 2, belaruskali: 1, amkodor: 1 },
    orders: [{
      id: "AI-1042",
      asset: "minfin",
      side: "Покупка",
      quantity: 2,
      total: 1190.19,
      price: 594.5,
      type: "market",
      typeLabel: "Рыночная",
      status: "Исполнена",
      createdAt: Date.now() - 86400000
    }],
    homeBlocks: { chart: true, favorites: true, starter: true },
    priceAlerts: true,
    orderAlerts: true,
    marketDigest: true,
    proMode: false
  },
  read() {
    try {
      return { ...structuredClone(this.initial), ...JSON.parse(localStorage.getItem(this.key) || "{}") };
    } catch {
      return structuredClone(this.initial);
    }
  },
  write(state) {
    localStorage.setItem(this.key, JSON.stringify(state));
  },
  reset() {
    localStorage.removeItem(this.key);
  },
  toggleFavorite(id) {
    const state = this.read();
    state.favorites = state.favorites.includes(id)
      ? state.favorites.filter(item => item !== id)
      : [...state.favorites, id];
    this.write(state);
    return state;
  },
  setProMode(enabled) {
    const state = this.read();
    state.proMode = !!enabled;
    this.write(state);
    return state;
  },
  saveOnboarding({ goal, experience, startAmount }) {
    const state = this.read();
    state.onboarded = true;
    state.goal = goal;
    state.experience = experience;
    state.startAmount = startAmount;
    this.write(state);
    return state;
  },
  commission(subtotal) {
    return subtotal * InvestData.commissionRate;
  },
  placeOrder({
    asset,
    side,
    quantity,
    type = "market",
    limitPrice = null,
    triggerPrice = null,
    trailPercent = null,
    timeInForce = "day"
  }) {
    const state = this.read();
    const item = InvestData.assetById(asset);
    if (!item) return { error: "Инструмент не найден" };

    const qty = Math.max(1, Math.floor(+quantity || 1));
    const typeMeta = InvestData.orderTypes.find(entry => entry.id === type) || InvestData.orderTypes[0];
    const isBuy = side === "Покупка";
    const execPrice = type === "market"
      ? item.price
      : +(limitPrice || triggerPrice || item.price);
    const subtotal = execPrice * qty;
    const commission = this.commission(subtotal);
    const total = isBuy ? subtotal + commission : subtotal - commission;

    if (isBuy && type === "market" && total > state.cash + 1e-9) {
      return { error: `Недостаточно средств. Доступно ${InvestData.format(state.cash)} BYN` };
    }
    if (!isBuy) {
      const held = state.holdings[asset] || 0;
      if (qty > held) {
        return { error: `Недостаточно бумаг. В портфеле ${held} шт.` };
      }
    }

    const immediate = type === "market";
    const order = {
      id: `AI-${Date.now().toString().slice(-5)}`,
      asset,
      side,
      quantity: qty,
      price: execPrice,
      limitPrice: limitPrice == null ? null : +limitPrice,
      triggerPrice: triggerPrice == null ? null : +triggerPrice,
      trailPercent: trailPercent == null ? null : +trailPercent,
      timeInForce,
      total,
      commission,
      type,
      typeLabel: typeMeta.label,
      status: immediate ? "Исполнена" : "Ожидает исполнения",
      createdAt: Date.now()
    };

    if (immediate) {
      if (isBuy) {
        state.cash -= total;
        state.holdings[asset] = (state.holdings[asset] || 0) + qty;
      } else {
        state.cash += total;
        state.holdings[asset] = Math.max(0, (state.holdings[asset] || 0) - qty);
        if (!state.holdings[asset]) delete state.holdings[asset];
      }
    }

    state.orders = [order, ...state.orders];
    this.write(state);
    return { order, subtotal, commission, total, immediate };
  },
  cancelOrder(id) {
    const state = this.read();
    const target = state.orders.find(order => order.id === id);
    if (!target || target.status !== "Ожидает исполнения") {
      return { error: "Отменить можно только ожидающую заявку" };
    }
    state.orders = state.orders.map(order =>
      order.id === id ? { ...order, status: "Отменена" } : order
    );
    this.write(state);
    return { order: { ...target, status: "Отменена" } };
  }
};
