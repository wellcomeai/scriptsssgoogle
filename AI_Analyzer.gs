// ============================================
// 🤖 КОНФИГУРАЦИЯ API
// ============================================

const CONFIG = {
  openai: {
    apiKey: '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',
    apiUrl: 'https://api.openai.com/v1/chat/completions',
    model: 'gpt-4o-mini'
  },
  telegram: {
    botToken: '8431768082:AAH_Uxug5b0Q4TdrUQMrMSMjChPzIwFy6Is',
    chatId: '418838097'
  },
  marker: '🤖 Обработано от ИИ'
};

// ============================================
// 🎯 ГЛАВНАЯ ФУНКЦИЯ АНАЛИЗА ДИАЛОГОВ
// ============================================

function анализироватьДиалоги() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetsToProcess = ['Входящие', 'Исходящие'];
    
    let totalProcessed = 0;
    
    for (let sheetName of sheetsToProcess) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) continue;
      
      Logger.log(`Обработка листа: ${sheetName}`);
      
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) continue;
      
      // Проходим по всем строкам
      for (let row = 2; row <= lastRow; row++) {
        const phone = sheet.getRange(row, 1).getValue();
        if (!phone || phone.toString().trim() === '') continue;
        
        // Получаем количество колонок с данными
        const lastCol = sheet.getLastColumn();
        
        // Проходим по всем колонкам с диалогами (начиная с B)
        for (let col = 2; col <= lastCol; col++) {
          const cell = sheet.getRange(row, col);
          const richTextValue = cell.getRichTextValue();
          
          if (!richTextValue) continue;
          
          const cellText = richTextValue.getText();
          if (cellText.trim() === '') continue;
          
          // Проверяем, не обработан ли уже
          if (cellText.includes(CONFIG.marker)) {
            continue;
          }
          
          Logger.log(`Найден необработанный диалог: ${sheetName}, строка ${row}, колонка ${col}`);
          
          // Обрабатываем диалог
          const processed = обработатьОдинДиалог(sheet, row, col, phone, cellText, richTextValue, sheetName);
          if (processed) {
            totalProcessed++;
          }
        }
      }
    }
    
    Logger.log(`Всего обработано диалогов: ${totalProcessed}`);
    
  } catch (error) {
    Logger.log('ОШИБКА в анализироватьДиалоги: ' + error.toString());
    Logger.log('Стек: ' + error.stack);
  }
}

// ============================================
// 📝 ОБРАБОТКА ОДНОГО ДИАЛОГА
// ============================================

function обработатьОдинДиалог(sheet, row, col, phone, dialogText, originalRichText, sheetType) {
  try {
    // Убираем эмодзи и временные метки для чистого текста
    const cleanDialog = очиститьТекстДиалога(dialogText);
    
    Logger.log(`Отправка в OpenAI...`);
    
    // Получаем анализ от OpenAI
    const analysis = отправитьВOpenAI(cleanDialog);
    
    if (!analysis) {
      Logger.log('Не удалось получить анализ от OpenAI');
      return false;
    }
    
    Logger.log(`Получен анализ: ${JSON.stringify(analysis)}`);
    
    // Формируем текст анализа
    const analysisText = форматироватьАнализ(analysis);
    
    // ✅ СОЗДАЕМ Rich Text с АНАЛИЗОМ В НАЧАЛЕ
    const richText = создатьФорматированныйАнализВНачале(originalRichText, analysisText);
    
    // Обновляем ячейку
    sheet.getRange(row, col).setRichTextValue(richText);
    
    Logger.log('Анализ добавлен в таблицу');
    
    // 🎨 ОБНОВЛЯЕМ ЦВЕТ РАМКИ ПО СТАТУСУ AI
    updateCellBorderByStatus(sheet, row, col, analysis.status);
    
    Logger.log(`Цвет рамки обновлен для статуса: ${analysis.status}`);
    
    // Отправляем в Telegram
    отправитьВTelegram(phone, analysis, sheetType);
    
    Logger.log('Уведомление отправлено в Telegram');
    
    return true;
    
  } catch (error) {
    Logger.log('ОШИБКА при обработке диалога: ' + error.toString());
    return false;
  }
}

// ============================================
// 🧹 ОЧИСТКА ТЕКСТА ДИАЛОГА
// ============================================

function очиститьТекстДиалога(text) {
  // Убираем только маркер обработки, оставляем остальное для контекста
  let cleaned = text.replace(CONFIG.marker, '');
  
  // Убираем множественные переносы строк
  cleaned = cleaned.replace(/\n{3,}/g, '\n\n');
  
  return cleaned.trim();
}

// ============================================
// 🤖 ОТПРАВКА ЗАПРОСА В OPENAI
// ============================================

function отправитьВOpenAI(dialogText) {
  try {
    const systemPrompt = `Ты менеджер по продажам. Проанализируй диалог между клиентом и ботом/менеджером.

Верни ТОЛЬКО валидный JSON в таком формате:
{
  "city": "город клиента или 'Не указан'",
  "name": "ФИО клиента или 'Не представился'",
  "interest": "Высокий|Средний|Низкий",
  "status": "Новый лид|Требует уточнений|Готов к сделке|Отказ",
  "summary": "2-3 предложения о чем говорили и какие договоренности"
}

Правила:
- Если город упомянут - укажи его
- Если клиент представился - укажи имя
- Оцени реальный интерес по содержанию диалога
- Статус определи по готовности клиента
- Резюме должно быть информативным для менеджера`;

    const payload = {
      model: CONFIG.openai.model,
      messages: [
        {
          role: 'system',
          content: systemPrompt
        },
        {
          role: 'user',
          content: dialogText
        }
      ],
      temperature: 0.3,
      response_format: { type: 'json_object' }
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + CONFIG.openai.apiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(CONFIG.openai.apiUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log('OpenAI API Error: ' + responseCode);
      Logger.log('Response: ' + response.getContentText());
      return null;
    }
    
    const result = JSON.parse(response.getContentText());
    const content = result.choices[0].message.content;
    
    return JSON.parse(content);
    
  } catch (error) {
    Logger.log('ОШИБКА при запросе к OpenAI: ' + error.toString());
    return null;
  }
}

// ============================================
// 📊 ФОРМАТИРОВАНИЕ АНАЛИЗА ДЛЯ ТАБЛИЦЫ
// ============================================

function форматироватьАнализ(analysis) {
  const divider = '━━━━━━━━━━━━━━━━━━━━';
  
  let text = divider + '\n';
  text += '🤖 АНАЛИЗ ОТ ИИ\n\n';
  text += `🏙️ Город: ${analysis.city}\n`;
  text += `👤 ФИО: ${analysis.name}\n`;
  text += `💼 Интерес: ${analysis.interest}\n`;
  text += `📊 Статус: ${analysis.status}\n\n`;
  text += `📝 Краткое резюме:\n${analysis.summary}\n`;
  text += divider;
  
  return text;
}

// ============================================
// 🎨 СОЗДАНИЕ RICH TEXT С АНАЛИЗОМ В НАЧАЛЕ
// ✨ НОВАЯ ЛОГИКА: Анализ ПЕРЕД диалогом
// ============================================

function создатьФорматированныйАнализВНачале(originalRichText, analysisText) {
  const originalText = originalRichText.getText();
  
  // Извлекаем дату и время из начала диалога
  const lines = originalText.split('\n');
  const dateTimeLine = lines[0]; // Первая строка с датой
  
  // Извлекаем текст диалога без даты
  const dialogWithoutDate = lines.slice(1).join('\n').trim();
  
  // ✅ НОВЫЙ ПОРЯДОК: Дата → Анализ → Диалог → Маркер
  const fullText = `${dateTimeLine}\n\n${analysisText}\n\n${dialogWithoutDate}\n\n${CONFIG.marker}`;
  
  const richTextBuilder = SpreadsheetApp.newRichTextValue().setText(fullText);
  
  let currentPos = 0;
  
  // 1. Форматируем дату (первая строка)
  richTextBuilder.setTextStyle(
    0,
    dateTimeLine.length,
    SpreadsheetApp.newTextStyle()
      .setBold(true)
      .setForegroundColor('#1a73e8')
      .setFontSize(10)
      .build()
  );
  currentPos = dateTimeLine.length + 2; // +2 для \n\n
  
  // 2. Форматируем блок анализа
  const analysisStart = currentPos;
  const analysisEnd = analysisStart + analysisText.length;
  
  const analysisLines = analysisText.split('\n');
  let analysisPos = analysisStart;
  
  for (let line of analysisLines) {
    if (line.startsWith('━')) {
      // Разделитель - серый
      richTextBuilder.setTextStyle(
        analysisPos,
        analysisPos + line.length,
        SpreadsheetApp.newTextStyle()
          .setForegroundColor('#9e9e9e')
          .setBold(true)
          .setFontSize(10)
          .build()
      );
    } else if (line.startsWith('🤖 АНАЛИЗ')) {
      // Заголовок - синий жирный
      richTextBuilder.setTextStyle(
        analysisPos,
        analysisPos + line.length,
        SpreadsheetApp.newTextStyle()
          .setForegroundColor('#1a73e8')
          .setBold(true)
          .setFontSize(11)
          .build()
      );
    } else if (line.startsWith('🏙️') || line.startsWith('👤') || line.startsWith('💼') || line.startsWith('📊')) {
      // Поля данных - жирный эмодзи + обычный текст
      const colonIndex = line.indexOf(':');
      if (colonIndex > 0) {
        richTextBuilder.setTextStyle(
          analysisPos,
          analysisPos + colonIndex + 1,
          SpreadsheetApp.newTextStyle()
            .setBold(true)
            .setForegroundColor('#202124')
            .setFontSize(10)
            .build()
        );
        richTextBuilder.setTextStyle(
          analysisPos + colonIndex + 1,
          analysisPos + line.length,
          SpreadsheetApp.newTextStyle()
            .setBold(false)
            .setForegroundColor('#5f6368')
            .setFontSize(10)
            .build()
        );
      }
    } else if (line.startsWith('📝')) {
      // Резюме - жирный заголовок
      richTextBuilder.setTextStyle(
        analysisPos,
        analysisPos + line.length,
        SpreadsheetApp.newTextStyle()
          .setBold(true)
          .setForegroundColor('#202124')
          .setFontSize(10)
          .build()
      );
    } else if (line.trim() !== '' && !line.startsWith('━')) {
      // Обычный текст резюме
      richTextBuilder.setTextStyle(
        analysisPos,
        analysisPos + line.length,
        SpreadsheetApp.newTextStyle()
          .setBold(false)
          .setForegroundColor('#5f6368')
          .setFontSize(10)
          .build()
      );
    }
    analysisPos += line.length + 1;
  }
  
  // 3. Форматируем диалог
  const dialogStart = analysisEnd + 2;
  const dialogLines = dialogWithoutDate.split('\n');
  let dialogPos = dialogStart;
  
  for (let line of dialogLines) {
    if (line.startsWith('👤 Клиент:')) {
      // Клиент - красный жирный для метки
      richTextBuilder.setTextStyle(
        dialogPos,
        dialogPos + 10,
        SpreadsheetApp.newTextStyle()
          .setBold(true)
          .setForegroundColor('#ea4335')
          .setFontSize(10)
          .build()
      );
      richTextBuilder.setTextStyle(
        dialogPos + 10,
        dialogPos + line.length,
        SpreadsheetApp.newTextStyle()
          .setBold(false)
          .setForegroundColor('#444444')
          .setFontSize(10)
          .build()
      );
    } else if (line.startsWith('🤖 Бот:')) {
      // Бот - зеленый жирный для метки
      richTextBuilder.setTextStyle(
        dialogPos,
        dialogPos + 7,
        SpreadsheetApp.newTextStyle()
          .setBold(true)
          .setForegroundColor('#34a853')
          .setFontSize(10)
          .build()
      );
      richTextBuilder.setTextStyle(
        dialogPos + 7,
        dialogPos + line.length,
        SpreadsheetApp.newTextStyle()
          .setBold(false)
          .setForegroundColor('#555555')
          .setFontSize(10)
          .build()
      );
    }
    dialogPos += line.length + 1;
  }
  
  // 4. Форматируем маркер обработки (зеленый жирный)
  const markerStart = fullText.length - CONFIG.marker.length;
  richTextBuilder.setTextStyle(
    markerStart,
    fullText.length,
    SpreadsheetApp.newTextStyle()
      .setForegroundColor('#34a853')
      .setBold(true)
      .setFontSize(9)
      .build()
  );
  
  return richTextBuilder.build();
}

// ============================================
// 📱 ОТПРАВКА УВЕДОМЛЕНИЯ В TELEGRAM
// ============================================

function отправитьВTelegram(phone, analysis, sheetType) {
  try {
    const now = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
    
    const callTypeEmoji = sheetType === 'Входящие' ? '📞' : '📱';
    const callTypeText = sheetType === 'Входящие' ? 'Входящий' : 'Исходящий';
    
    let message = `🔔 НОВАЯ ЗАЯВКА (${callTypeText})\n\n`;
    message += `${callTypeEmoji} Телефон: +${phone}\n`;
    message += `🏙️ Город: ${analysis.city}\n`;
    message += `👤 ФИО: ${analysis.name}\n`;
    message += `💼 Интерес: ${analysis.interest}\n`;
    message += `📊 Статус: ${analysis.status}\n\n`;
    message += `📝 Резюме:\n${analysis.summary}\n\n`;
    message += `🕐 ${dateStr}`;
    
    const telegramUrl = `https://api.telegram.org/bot${CONFIG.telegram.botToken}/sendMessage`;
    
    const payload = {
      chat_id: CONFIG.telegram.chatId,
      text: message,
      parse_mode: 'HTML'
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(telegramUrl, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log('Telegram API Error: ' + responseCode);
      Logger.log('Response: ' + response.getContentText());
    } else {
      Logger.log('Telegram: Сообщение успешно отправлено');
    }
    
  } catch (error) {
    Logger.log('ОШИБКА при отправке в Telegram: ' + error.toString());
  }
}

// ============================================
// ⏰ УПРАВЛЕНИЕ ТРИГГЕРАМИ ДЛЯ АНАЛИЗА
// ============================================

function установитьАвтоанализ() {
  try {
    удалитьАвтоанализ();
    
    ScriptApp.newTrigger('анализироватьДиалоги')
      .timeBased()
      .everyMinutes(1)
      .create();
    
    SpreadsheetApp.getUi().alert('✅ Автоанализ включен!\n\nСкрипт будет проверять новые диалоги каждую минуту.');
    
    onOpen();
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Ошибка: ' + error.toString());
  }
}

function удалитьАвтоанализ() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deletedCount = 0;
    
    for (let trigger of triggers) {
      if (trigger.getHandlerFunction() === 'анализироватьДиалоги') {
        ScriptApp.deleteTrigger(trigger);
        deletedCount++;
      }
    }
    
    if (deletedCount > 0) {
      SpreadsheetApp.getUi().alert('✅ Автоанализ выключен');
    } else {
      SpreadsheetApp.getUi().alert('ℹ️ Автоанализ не был включен');
    }
    
    onOpen();
    
  } catch (error) {
    SpreadsheetApp.getUi().alert('❌ Ошибка: ' + error.toString());
  }
}

function переключитьАвтоанализ() {
  const triggers = ScriptApp.getProjectTriggers();
  let isActive = false;
  
  for (let trigger of triggers) {
    if (trigger.getHandlerFunction() === 'анализироватьДиалоги') {
      isActive = true;
      break;
    }
  }
  
  if (isActive) {
    удалитьАвтоанализ();
  } else {
    установитьАвтоанализ();
  }
}

// ============================================
// 🧪 ТЕСТ TELEGRAM
// ============================================

function тестТелеграм() {
  try {
    Logger.log('Начало теста Telegram...');
    Logger.log('Bot Token: ' + CONFIG.telegram.botToken.substring(0, 10) + '...');
    Logger.log('Chat ID: ' + CONFIG.telegram.chatId);
    
    const telegramUrl = `https://api.telegram.org/bot${CONFIG.telegram.botToken}/sendMessage`;
    
    const payload = {
      chat_id: CONFIG.telegram.chatId,
      text: '🧪 Тест связи!\n\nЕсли видишь это сообщение - всё работает! ✅'
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(telegramUrl, options);
    const responseCode = response.getResponseCode();
    
    Logger.log('Response Code: ' + responseCode);
    Logger.log('Response: ' + response.getContentText());
    
    if (responseCode === 200) {
      SpreadsheetApp.getUi().alert('✅ Telegram работает!\n\nПроверьте сообщение в боте.');
    } else {
      SpreadsheetApp.getUi().alert('❌ Ошибка Telegram:\n\n' + response.getContentText());
    }
    
  } catch (error) {
    Logger.log('ОШИБКА в тесте: ' + error.toString());
    SpreadsheetApp.getUi().alert('❌ Ошибка: ' + error.toString());
  }
}
