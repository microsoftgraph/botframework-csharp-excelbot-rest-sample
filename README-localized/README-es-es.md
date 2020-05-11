---
page_type: sample
products:
- office-excel
- ms-graph
languages:
- csharp
description: "El bot de Excel es un bot creado con Microsoft Bot Framework que muestra cómo usar Excel con la API de Microsoft Graph"
extensions:
  contentType: samples 
  technologies:
  - Microsoft Graph
  - Microsoft Bot Framework
  services:
  - Excel
  createdDate: 9/15/2016 10:30:08 PM
---

# Bot de Excel

## Tabla de contenido ##

[Introducción](#introduction)

[Requisitos previos](#prerequisites)

[Clonar o descargar el repositorio](#Cloning-or-downloading-this-repository)

[Configurar el inquilino de Azure AD](#Configure-your-Azure-AD-tenant)

[Registrar el bot](#Register-the-bot)

[Envíenos sus comentarios](#Give-us-your-feedback)

## Introducción
<a name="introduction"></a>
 El bot de Excel es un ejemplo en el que se demuestra cómo usar [Microsoft Graph](https://graph.microsoft.io) y específicamente la [API de REST de Excel](https://graph.microsoft.io/es-es/docs/api-reference/v1.0/resources/excel) para tener acceso a los libros de Excel almacenados en OneDrive para la Empresa mediante una interfaz de usuario de conversación. Esta escrito en C# y usa [Microsoft Bot Framework](https://dev.botframework.com/) y [Language Understanding Intelligent Service (LUIS)](https://www.luis.ai/).

*Nota*: El código de este ejemplo se ha escrito originalmente para un prototipo de experiencia de usuario y no muestra necesariamente cómo crear un código con calidad de producción.

## Requisitos previos
<a name="prerequisites"></a>

Para este ejemplo se necesita lo siguiente:  

- Visual Studio 2017.
- Una cuenta de Office 365 para empresas. Puede registrarse para obtener una [suscripción de Office 365 Developer](https://msdn.microsoft.com/es-es/office/office365/howto/setup-development-environment) que incluye los recursos que necesita para comenzar a crear aplicaciones de Office 365.

## Clonar o descargar el repositorio
<a name="cloning-downloading-repo"></a>

- Clone este repositorio en una carpeta local.

    ` git clone https://github.com/nicolesigei/botframework-csharp-excelbot-rest-sample.git `

<a name="configure-azure"></a>
## Configurar el inquilino de Azure AD

1. Abra el explorador y vaya al [Centro de administración de Azure Active Directory](https://aad.portal.azure.com). Inicie sesión con **una cuenta profesional o educativa**.

1. Seleccione **Azure Active Directory** en el panel de navegación izquierdo y, después seleccione**registros de aplicaciones** en **Administrar**.

    ![Captura de pantalla de los registros de la aplicación ](readme-images/aad-portal-app-registrations.png)

1. Seleccione **Nuevo registro**. En la página **Registrar una aplicación**, establezca los valores siguientes.

    - Establezca un **Nombre** de preferencia; por ejemplo, `Aplicación de bot de Excel`.
    - Establezca los **Tipos de cuenta compatibles** en **Cuentas de cualquier directorio organizativo**.
    - En **URI de redirección**, establezca la primera lista desplegable en `Web` y el valor en http://localhost:3978/callback.

    ![Captura de pantalla de la página Registrar una aplicación](readme-images/aad-register-an-app.PNG)

    > **Nota:** Si está ejecutando de forma local y en Azure, debe agregar dos URL de redirección aquí, una a la instancia local y otra a la aplicación web de Azure.
    
1. Haga clic en **Registrar**. En la página de **Aplicación de bot de Excel**, copie el valor de **Id. de aplicación (cliente)** y guárdelo, lo necesitará para configurar la aplicación.

    ![Captura de pantalla del Id. de aplicación del nuevo registro](readme-images/aad-application-id.PNG)

1. Seleccione **Certificados y secretos** en **Administrar**. Seleccione el botón **Nuevo secreto de cliente**. Escriba un valor en **Descripción** y seleccione una de las opciones de **Expirar** y luego seleccione **Agregar**.

    ![Captura de pantalla del diálogo Agregar un cliente secreto](readme-images/aad-new-client-secret.png)

1. Copie el valor del secreto de cliente antes de salir de esta página. Lo necesitará para configurar la aplicación.

    > [¡IMPORTANTE!]
    > El secreto de cliente no se mostrará otra vez, asegúrese de copiarlo en este momento.

    ![Captura de pantalla del nuevo secreto de cliente agregado](readme-images/aad-copy-client-secret.png)
	<a name = "register-bot"></a>
## Registrar el bot

Realice estos pasos para configurar el entorno de desarrollo a fin de crear y probar el bot de Excel:

- Descargue e instale [Azure Cosmos DB Emulator](https://docs.microsoft.com/es-es/azure/cosmos-db/local-emulator).

- Realice una copia de **./ExcelBot/PrivateSettings.config.example** en el mismo directorio. Cambie el nombre del archivo a **PrivateSettings.config**.
- Abra el archivo de solución ExcelBot.sln.
- Registre el bot en [Bot Framework](https://dev.botframework.com/bots/new).
- Copie los valores de MicrosoftAppId y MicrosoftAppPassword del bot al archivo PrivateSettings.config
- Registre el bot para llamar a Microsoft Graph.
- Copie el **Id. de cliente** y el **secreto** de Azure Active Directory al archivo PrivateSettings.config.
- Cree un nuevo modelo en el servicio [LUIS](https://www.luis.ai).
- Importe el archivo LUIS\\excelbot.json a LUIS.
- Entrene y publique el modelo de LUIS.
- Copie el Id. de modelo y la clave de suscripción de LUIS al archivo Dialogs\\ExcelBotDialog.cs.
- (Opcional) Habilite el chat web para el bot en Bot Framework y copie la plantilla de inserción de web chat en el archivo chat.htm
- (Opcional) Para que el bot envíe la telemetría a [Visual Studio Application Insights](https://azure.microsoft.com/es-es/services/application-insights/), copie la clave de instrumentación en los siguientes archivos: ApplicationInsights.config, default.htm, loggedin.htm, chat.htm
- Compile la solución.
- Presione F5 para iniciar el bot de forma local.
- Pruebe el bot de forma local mediante [Bot Framework Emulator](https://docs.botframework.com/es-es/tools/bot-framework-emulator).
- Cree una base de datos de Azure Cosmos DB en Azure que use la API de SQL.
- Reemplace el nombre de host del bot en el archivo PrivateSettings.config.
- Reemplace el URI y la clave de la base datos en el archivo PrivateSettings.config.
- Publique la solución en la aplicación web de Azure.
- Navegue a la página chat.htm para probar el bot implementado mediante el control de web chat.  

## Envíenos sus comentarios

<a name="Give-us-your-feedback"></a>

Su opinión es importante para nosotros.  

Revise el código de ejemplo y háganos saber todas las preguntas y las dificultades que encuentre [enviando un problema](https://github.com/microsoftgraph/botframework-csharp-excelbot-rest-sample/issues) directamente en este repositorio. Incluya los pasos de reproducción, las salidas de la consola y los mensajes de error en cualquier problema que envíe.

Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.

## Copyright

Copyright (c) 2019 Microsoft. Todos los derechos reservados.
  
