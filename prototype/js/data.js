window.InvestData = {
  commissionRate: 0.001,
  assets: [
    { id: "minfin", ticker: "МФ", name: "Облигации Минфина", type: "Облигация", sector: "Госсектор", price: 594.5, change: 0.9, yield: "8,5% годовых", risk: "Низкий", recommended: true, top: true },
    { id: "belaruskali", ticker: "БК", name: "Беларуськалий", type: "Акция", sector: "Промышленность", price: 1486, change: 3.1, yield: "—", risk: "Средний", top: true },
    { id: "amkodor", ticker: "АМ", name: "Амкодор", type: "Акция", sector: "Промышленность", price: 168.2, change: -0.6, yield: "—", risk: "Средний", top: true },
    { id: "priorbank", ticker: "ПБ", name: "Приорбанк", type: "Акция", sector: "Финансы", price: 92.4, change: 1.2, yield: "6,2% годовых", risk: "Средний", top: true },
    { id: "mtz", ticker: "МТЗ", name: "Облигации МТЗ", type: "Облигация", sector: "Промышленность", price: 1020, change: 0.3, yield: "9,1% годовых", risk: "Низкий", top: true },
    { id: "bps", ticker: "БПС", name: "БПС-Сбербанк", type: "Акция", sector: "Финансы", price: 78.1, change: -0.4, yield: "—", risk: "Средний" },
    { id: "belinvest", ticker: "БИ", name: "Белинвестбанк", type: "Акция", sector: "Финансы", price: 54.3, change: 0.7, yield: "4,8% годовых", risk: "Низкий" },
    { id: "mtel", ticker: "МТС", name: "МТС Беларусь", type: "Акция", sector: "IT и связь", price: 210.0, change: 1.8, yield: "5,1% годовых", risk: "Средний" }
  ],
  goals: {
    "Накопить на цель": {
      title: "Стартовый набор под цель",
      hint: "Больше облигаций для стабильности и немного акций роста",
      ids: ["minfin", "mtz", "belaruskali"]
    },
    "Получать пассивный доход": {
      title: "Набор для пассивного дохода",
      hint: "Купоны и дивиденды с умеренным риском",
      ids: ["minfin", "priorbank", "mtz"]
    },
    "Просто попробовать": {
      title: "Простой старт",
      hint: "Небольшая сумма, понятные бумаги и низкий порог входа",
      ids: ["minfin", "amkodor", "priorbank"]
    }
  },
  orderTypes: [
    { id: "market", label: "Рыночная", hint: "Исполняется сразу по лучшей доступной цене" },
    { id: "limit", label: "Лимитная", hint: "Исполняется только по вашей цене или лучше" },
    { id: "stop", label: "Стоп-лосс", hint: "Продажа при падении цены до триггера" },
    { id: "take", label: "Тейк-профит", hint: "Фиксация прибыли при росте до цели" },
    { id: "trailing", label: "Trailing stop", hint: "Стоп следует за ценой с заданным отступом" }
  ],
  terms: [
    { term: "Стакан", text: "Список заявок на покупку и продажу по разным ценам. Показывает, где сейчас спрос и предложение." },
    { term: "Рыночная заявка", text: "Покупка или продажа по текущей рыночной цене. Исполняется сразу." },
    { term: "Лимитная заявка", text: "Вы задаёте максимальную цену покупки или минимальную цену продажи." },
    { term: "Стоп-лосс", text: "Защитная заявка: при падении цены до уровня срабатывает продажа." },
    { term: "Тейк-профит", text: "Заявка на фиксацию прибыли, когда цена достигает целевого уровня." },
    { term: "Trailing stop", text: "Стоп, который автоматически подтягивается вслед за ростом цены." },
    { term: "Спред", text: "Разница между лучшей ценой покупки и лучшей ценой продажи." },
    { term: "Комиссия", text: "Плата брокеру за сделку. В Alfa Invest она всегда видна до подтверждения." }
  ],
  stories: ["Рынок", "Идеи", "Курс", "Новости", "Альфа"],
  format(value) {
    return new Intl.NumberFormat("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
  },
  formatQty(value) {
    return new Intl.NumberFormat("ru-RU").format(Math.round(value));
  },
  orderBook(asset) {
    const step = asset.price >= 500 ? 0.5 : asset.price >= 100 ? 0.2 : 0.1;
    const mid = Math.round(asset.price / step) * step;
    const asks = [];
    const bids = [];
    const volumes = [3125, 6729, 9808, 4520, 12880, 2340, 8150, 3960, 10420, 5610];
    for (let i = 0; i < 8; i++) {
      asks.push({ price: +(mid + step * (8 - i)).toFixed(2), quantity: volumes[i] + i * 180 });
      bids.push({ price: +(mid - step * (i + 1)).toFixed(2), quantity: volumes[9 - i] + i * 210 });
    }
    const bestAsk = asks[asks.length - 1].price;
    const bestBid = bids[0].price;
    return {
      asks,
      bids,
      bestAsk,
      bestBid,
      spread: +(bestAsk - bestBid).toFixed(2),
      maxQty: Math.max(...asks.map(level => level.quantity), ...bids.map(level => level.quantity))
    };
  },
  assetById(id) {
    return this.assets.find(asset => asset.id === id);
  }
};
