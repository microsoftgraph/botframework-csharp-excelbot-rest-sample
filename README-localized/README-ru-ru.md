---
page_type: sample
products:
- office-excel
- ms-graph
languages:
- csharp
description: "Excel Bot - это бот, созданный на основе Microsoft Bot Framework, который демонстрирует, как использовать Excel с Microsoft Graph API"
extensions:
  contentType: samples 
  technologies:
  - Microsoft Graph
  - Microsoft Bot Framework
  services:
  - Excel
  createdDate: 9/15/2016 10:30:08 PM
---

# Excel Bot

## Содержание ##

[Вступление.](#introduction)

[Требования.](#prerequisites)

[Клонирование или скачивание этого репозитория.](#Cloning-or-downloading-this-repository)

[Настройка клиента Azure AD.](#Configure-your-Azure-AD-tenant)

[Зарегистрировать Bot.](#Register-the-bot)

[Оставьте свой отзыв](#Give-us-your-feedback)

## Вступление.
<a name="introduction"></a>
Excel Bot — это пример, в котором показано, как использовать [Microsoft Graph](https://graph.microsoft.io), а именно [Excel REST API](https://graph.microsoft.io/en-us/docs/api-reference/v1.0/resources/excel) для доступа к книгам Excel, хранящимся в службе OneDrive для бизнеса с помощью диалогового пользовательского интерфейса. Он написан на C # и использует [Microsoft Bot Framework](https://dev.botframework.com/) и [интеллектуальную службу понимания языка (LUIS)](https://www.luis.ai/).

*Примечание*. Код в этом примере изначально был написан для прототипа пользовательского интерфейса и не обязательно демонстрирует, как создать производственный качественный код.

## Требования.
<a name="prerequisites"></a>

Для этого примера требуются приведенные ниже компоненты.  

- Visual Studio 2017.
- Учетная запись Office 365 для бизнеса. Вы можете зарегистрироваться для получения подписки на [Office 365 для разработчиков ](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment) с ресурсами, которые необходимо приступить к созданию приложений Office 365.

## Клонирование или скачивание этого репозитория.
<a name="cloning-downloading-repo"></a>

- Клонирование этого репозитория в локальную папку

    ` git clone https://github.com/nicolesigei/botframework-csharp-excelbot-rest-sample.git `

<a name="configure-azure"></a>
## Настройка клиента Azure AD.

1. Откройте браузер и перейдите в [Центр администрирования Azure Active Directory](https://aad.portal.azure.com). Войдите с помощью **рабочей или учебной учетной записи**.

1. Выберите **Azure Active Directory** на панели навигации слева, а затем — пункт **Регистрация приложения** в разделе **Управление**.

    ![Снимок экрана: пункт "Регистрация приложения"](readme-images/aad-portal-app-registrations.png)

1. Выберите **Новая регистрация**. На странице **Регистрация приложения** задайте необходимые значения следующим образом:

    - Установите предпочтительное **Имя**, например, `Приложение Excel Bot`.
    - Задайте для параметра **Поддерживаемые типы учетных записей** значение **Учетные записи в любом каталоге организации**.
    - В разделе **перенаправления URI** выберите в первом раскрывающемся списке значение `веб-` и задайте для него значение http://localhost:3978/callback.

    ![Снимок экрана: страница "Регистрация приложения"](readme-images/aad-register-an-app.PNG)

    > **Примечание.** Если вы используете это локально и в Azure, вам нужно добавить здесь два URL-адреса перенаправления, один для локального экземпляра и один для веб-приложения Azure.
    
1. Нажмите кнопку **Зарегистрировать**. На странице **приложения Excel Bot** скопируйте значение **идентификатора приложения (клиента)** и сохраните его, оно понадобится вам для настройки приложения.

    ![Снимок экрана: идентификатор приложения для новой регистрации](readme-images/aad-application-id.PNG)

1. Выберите **Сертификаты и секреты** в разделе **Управление**. Нажмите кнопку **Новый секрет клиента**. Введите значение в поле **Описание**, выберите один из вариантов для **Срок действия** и нажмите **Добавить**.

    ![Снимок экрана: диалоговое окно "Добавление секрета клиента"](readme-images/aad-new-client-secret.png)

1. Скопируйте значение секрета клиента, а затем покиньте эту страницу. Он понадобится вам для настройки приложения.

    > [ВАЖНО!]
    > Этот секрет клиента больше не будет отображаться, поэтому обязательно скопируйте его.

    ![Снимок экрана: только что добавленный секрет клиента](readme-images/aad-copy-client-secret.png)
	<a name = "register-bot"></a>
## Зарегистрировать Bot.

Выполните эти шаги, чтобы настроить среду разработки для сборки и тестирования бота Excel:

- Загрузите и установите [эмулятор Azure Cosmos DB](https://docs.microsoft.com/en-us/azure/cosmos-db/local-emulator)

- Сделайте копию **./ExcelBot/PrivateSettings.config.example** в том же каталоге. Назовите файл **PrivateSettings.config**.
- Откройте файл решения ExcelBot.sln
- Зарегистрируйте бота в [Bot Framework](https://dev.botframework.com/bots/new)
- Скопируйте бот MicrosoftAppId и MicrosoftAppPassword в файл PrivateSettings.config
- Зарегистрируйте бота, чтобы позвонить в Microsoft Graph.
- Скопируйте **идентификатор клиента** Azure Active Directory и **секретный файл** в файл PrivateSettings.config.
- Создание новой модели в службе [LUIS](https://www.luis.ai)
- Импортируйте файл LUIS \\ excelbot.json в LUIS
- Обучите и опубликуйте модель LUIS
- Скопируйте идентификатор модели LUIS и ключ подписки в файл Dialogs \\ ExcelBotDialog.cs.
- (Необязательно) Включите веб-чат для бота в Bot Framework и скопируйте шаблон для встраивания веб-чата в файл chat.htm
- Желательно Чтобы сделать так, чтобы в [приложении Visual Studio](https://azure.microsoft.com/en-us/services/application-insights/)Application Insights, скопируйте ключ инструментирования в следующие файлы: ApplicationInsights.config, default.htm, loggedin.htm, chat.htm
- Постройте решение
- Нажмите F5, чтобы запустить бот локально
- Протестируйте бот локально с [эмулятором Bot Framework](https://docs.botframework.com/en-us/tools/bot-framework-emulator)
- Создайте базу данных Azure Cosmos в Azure, которая использует SQL API
- Замените имя хоста ботов в файле PrivateSettings.config
- Замените URI базы данных и введите ключ в файл PrivateSettings.config.
- Опубликуйте решение в веб-приложении Azure.
- Протестируйте развернутого бота с помощью элемента управления веб-чата, перейдя на страницу chat.htm  

## Оставьте свой отзыв

<a name="Give-us-your-feedback"></a>

Ваш отзыв важен для нас.  

Ознакомьтесь с образцом кода и сообщите нам о любых возникших вопросах и проблемах, с которыми вы столкнулись, [отправив сообщение](https://github.com/microsoftgraph/botframework-csharp-excelbot-rest-sample/issues) в этом репозитории. Укажите выполненные действия, выходное сообщение консоли и сообщения об ошибках при любой проблеме.

Этот проект соответствует [Правилам поведения разработчиков открытого кода Майкрософт](https://opensource.microsoft.com/codeofconduct/). Дополнительные сведения см. в разделе [Часто задаваемые вопросы о правилах поведения](https://opensource.microsoft.com/codeofconduct/faq/). Если у вас возникли вопросы или замечания, напишите нам по адресу [opencode@microsoft.com](mailto:opencode@microsoft.com).

## Авторские права

(c) Корпорация Майкрософт (Microsoft Corporation), 2019. Все права защищены.
  
