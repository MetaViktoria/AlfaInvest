window.InvestData = {
  assets: [
    { id: "minfin", ticker: "МФ", name: "Облигации Минфина", type: "Облигация", sector: "Госсектор", price: 594.5, change: 0.9, yield: "8,5% годовых", risk: "Низкий", recommended: true },
    { id: "belaruskali", ticker: "БК", name: "Беларуськалий", type: "Акция", sector: "Промышленность", price: 1486, change: 3.1, yield: "—", risk: "Средний" },
    { id: "amkodor", ticker: "АМ", name: "Амкодор", type: "Акция", sector: "Промышленность", price: 168.2, change: -0.6, yield: "—", risk: "Средний" },
    { id: "priorbank", ticker: "ПБ", name: "Приорбанк", type: "Акция", sector: "Финансы", price: 92.4, change: 1.2, yield: "6,2% годовых", risk: "Средний" },
    { id: "mtz", ticker: "МТЗ", name: "Облигации МТЗ", type: "Облигация", sector: "Промышленность", price: 1020, change: 0.3, yield: "9,1% годовых", risk: "Низкий" }
  ],
  stories: ["Рынок", "Идеи", "Курс", "Новости", "Альфа"],
  format(value) {
    return new Intl.NumberFormat("ru-RU", { minimumFractionDigits: 2, maximumFractionDigits: 2 }).format(value);
  }
};
