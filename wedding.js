const { Builder, By } = require('selenium-webdriver');
const XLSX = require('xlsx');

// Функция для очистки текста
const cleanText = (text) => {
  if (!text) return '';
  return text
    .replace(/\s+/g, ' ')
    .replace(/\n/g, ' ')
    .trim();
};

// Функция перехода на следующую страницу
const goToNextPage = async (driver) => {
  try {
    const nextButton = await driver.findElement(By.css('.pagination-next'));
    const nextPageLink = await nextButton.getAttribute('href');
    if (nextPageLink) {
      await driver.get(nextPageLink);
      await driver.sleep(3000);
      return true;
    }
    return false;
  } catch (error) {
    console.error('Кнопка следующей страницы не найдена:', error.message);
    return false;
  }
};

// Основная функция
(async () => {
  const driver = await new Builder().forBrowser('safari').build();

  try {
    const maxPages = 5;
    let currentPage = 1;

    await driver.get('https://flowwow.com/moscow/wedding-cakes/');
    await driver.sleep(3000);

    const data = [];
    data.push([
      'Ссылка на товар',
      'Название',
      'Цена',
      'Валюта',
      'Рейтинг',
      'Количество отзывов',
      'Ссылка на изображение',
      'Стоимость доставки',
    ]);

    let hasNextPage = true;

    while (hasNextPage && currentPage <= maxPages) {
      console.log(`Обработка страницы ${currentPage}`);

      const products = await driver.findElements(By.css('article.tab-content-products-item'));

      for (let product of products) {
        try {
          const link = cleanText(await product.findElement(By.css('a.product-card')).getAttribute('href'));
          const name = cleanText(await product.findElement(By.css('.name')).getText());
          const price = cleanText(await product.findElement(By.css('.price span')).getText());
          const currency = cleanText(await product.findElement(By.css('.price i')).getText());
          const rating = cleanText(await product.findElement(By.css('.reviews-rating')).getText());
          const reviews = cleanText(await product.findElement(By.css('.reviews-count')).getText());
          const image = cleanText(await product.findElement(By.css('.product-photo img')).getAttribute('src'));
          const deliveryCost = cleanText(await product.findElement(By.css('.delivery-price span')).getText());

          data.push([link, name, price, currency, rating, reviews, image, deliveryCost]);
        } catch (err) {
          console.error('Ошибка при обработке товара:', err.message);
        }
      }

      currentPage++;
      if (currentPage <= maxPages) {
        hasNextPage = await goToNextPage(driver);
      } else {
        hasNextPage = false;
      }
    }

    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Товары');

    const filePath = 'flowwow-wedding-cakes.xlsx';
    XLSX.writeFile(workbook, filePath);

    console.log(`Таблица успешно создана: ${filePath}`);
  } catch (error) {
    console.error('Ошибка:', error.message);
  } finally {
    await driver.quit();
  }
})();
