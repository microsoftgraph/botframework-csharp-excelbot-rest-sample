---
page_type: sample
products:
- office-excel
- ms-graph
languages:
- csharp
description: "Excel 机器人是一个使用 Microsoft Bot Framework 生成的机器人，它演示了如何将 Excel 与 Microsoft Graph API 结合使用"
extensions:
  contentType: samples 
  technologies:
  - Microsoft Graph
  - Microsoft Bot Framework
  services:
  - Excel
  createdDate: 9/15/2016 10:30:08 PM
---

# Excel 机器人

## 目录。 ##

[简介。](#introduction)

[先决条件。](#prerequisites)

[复制或下载此存储库。](#Cloning-or-downloading-this-repository)

[配置 Azure AD 租户。](#Configure-your-Azure-AD-tenant)

[注册机器人。](#Register-the-bot)

[提供反馈](#Give-us-your-feedback)

## 简介。
<a name="introduction"></a>
Excel 机器人是一个示例，演示如何使用 [Microsoft Graph](https://graph.microsoft.io)、特别是 [ Excel REST API](https://graph.microsoft.io/en-us/docs/api-reference/v1.0/resources/excel) ，并通过会话用户界面访问存储在 OneDrive for Business 中的 Excel 工作簿。采用 C# 编写，使用了[Microsoft 机器人框架](https://dev.botframework.com/)和[语言理解智能服务(LUIS)](https://www.luis.ai/)。

*注意*：此示例中的代码最初专为用户体验原型编写，并不一定说明如何生成生产指令代码。

## 先决条件。
<a name="prerequisites"></a>

此示例要求如下：  

- Visual Studio 2017。
- Office 365 商业版帐户。你可以注册 [Office 365 开发人员订阅](https://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment)，其中包含你开始构建 Office 365 应用所需的资源。

## 复制或下载此存储库。
<a name="cloning-downloading-repo"></a>

- 复制此存储库至本地文件夹

    ` git clone https://github.com/nicolesigei/botframework-csharp-excelbot-rest-sample.git `

<a name="configure-azure"></a>
## 配置 Azure AD 租户。

1. 打开浏览器，并转到 [Azure Active Directory 管理中心](https://aad.portal.azure.com)。使用**工作或学校帐户**登录。

1. 选择左侧导航栏中的**Azure Active Directory**，再选择**管理**下的**应用注册**。

    ![“应用注册”的屏幕截图](readme-images/aad-portal-app-registrations.png)

1. 选择**“新注册”**。在“**注册应用**”页上，按如下方式设置值。

    - 设置首选**名称** ，如`Excel Bot App`。
    - 将“**支持的帐户类型**”设置为“**任何组织目录中的帐户**”。
    - 在“**重定向 URI**”下，将第一个下拉列表设置为“`Web`”，并将值设置为 http://localhost:3978/callback。

    ![“注册应用程序”页的屏幕截图](readme-images/aad-register-an-app.PNG)

    > **注意：**如果在本地和 Azure 上运行，应在此添加两个重定向 URL，一个至本地实例，一个至 Azure web 应用。
    
1. 选择“**注册**”。在“**Excel 机器人应用程序**”页面上，复制“**应用程序（客户端）ID**”值并保存，在配置应用程序时将会使用此数值。

    ![新应用注册的应用程序 ID 的屏幕截图](readme-images/aad-application-id.PNG)

1. 选择“**管理**”下的“**证书和密码**”。选择**新客户端密码**按钮。在**说明**中输入值，并选择一个**过期**选项，再选择**添加**。

    ![“添加客户端密码”对话框的屏幕截图](readme-images/aad-new-client-secret.png)

1. 离开此页前，先复制客户端密码值。你将需要它来配置应用。

    > [重要提示！]
    > 此客户端密码不会再次显示，所以请务必现在就复制它。

    ![新添加的客户端密码的屏幕截图](readme-images/aad-copy-client-secret.png)
	<a name = "register-bot"></a>
## 注册机器人。

完成下列步骤设置开发环境，以创建和测试 Excel 机器人：

- 下载并安装“[Azure Cosmos DB 模拟器](https://docs.microsoft.com/en-us/azure/cosmos-db/local-emulator)”

- 在同一目录中创建 **GraphWebHooks/PrivateSettings.example.config** 的副本。将文件命名为 **PrivateSettings.config**。
- 打开 ExcelBot.sln 解决方案文件
- 将机器人注册至“[机器人框架](https://dev.botframework.com/bots/new)”
- 将机器人 MicrosoftAppId 和 MicrosoftAppPassword 复制到 PrivateSettings.config 文件中
- 注册机器人以调用 Microsoft Graph。
- 复制 Azure Active Directory “**客户端 Id**”和“**密码**”至 PrivateSettings.config 文件
- 在 [LUIS](https://www.luis.ai) 服务中新建模型
- 导入 LUIS\\excelbot.json 文件至 LUIS 中
- 训练并发布 LUIS 模型
- 将 LUIS 模型 ID 和订阅密钥复制到 Dialogs\\ExcelBotDialog.cs 文件
- （可选）在机器人框架中启用机器人网络聊天并复制网络聊天嵌入模板至 chat.htm 文件
- （可选）若要获取机器人以发送遥测至 [Visual Studio Application Insights](https://azure.microsoft.com/en-us/services/application-insights/)，复制检测密钥至下列文件：ApplicationInsights.config, default.htm, loggedin.htm, chat.htm
- 构建解决方案
- 按下 F5 以本地启动机器人
- 使用[机器人框架模拟器](https://docs.botframework.com/en-us/tools/bot-framework-emulator)来本地测试机器人
- 在使用 SQL API 的 Azure 中创建 Azure Cosmos DB
- 在 PrivateSettings.config 文件中替换机器人主机名
- 在 PrivateSettings.config 文件中替换数据库 URI 和密钥
- 发布解决方案至 Azure web 应用
- 通过浏览至 chat.htm 页面使用网络聊天控件测试部署的机器人  

## 提供反馈

<a name="Give-us-your-feedback"></a>

我们非常重视你的反馈意见。  

查看示例代码并在此存储库中直接[提交问题](https://github.com/microsoftgraph/botframework-csharp-excelbot-rest-sample/issues)，告诉我们发现的任何疑问和问题。在任何打开话题中提供存储库步骤、控制台输出、错误消息。

此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则常见问题解答](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。

## 版权信息

版权所有 (c) 2019 Microsoft。保留所有权利。
  
