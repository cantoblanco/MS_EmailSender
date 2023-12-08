
# MS_EmailSender 项目

## 简介

EmailSender 是一个基于Python的应用程序，用于通过图形用户界面（GUI）发送电子邮件。该项目整合了PyQt5来实现GUI组件，提供了一个交互式的电子邮件编写和发送界面。它包括导入HTML电子邮件模板、编写带主题和内容的电子邮件以及发送给指定收件人的功能。

## 配置

在使用应用程序之前，请在`config.ini`中配置必要的细节。这包括设置电子邮件服务（如Microsoft Graph API）的认证细节，如`TOKEN_URL`、`CLIENT_ID`和`CLIENT_SECRET`。

## 使用方法

1. **启动应用程序**：运行`EmailSender.py`以启动应用程序。将打开标题为“白歌邮件发送器”的GUI窗口。

2. **编写电子邮件**：
   - 在“收件人”字段中输入收件人的电子邮件地址，用逗号分隔。
   - 在“主题”字段中提供电子邮件的主题。
   - 在“内容”部分编写或导入电子邮件内容。使用“导入HTML”按钮导入HTML文件作为电子邮件正文。

3. **发送电子邮件**：
   - 编写完电子邮件后，点击“发送邮件”按钮发送电子邮件。电子邮件发送的状态将显示在状态区域。

4. **导入电子邮件模板**：
   - 你可以导入HTML电子邮件模板（默认为`email.html`）作为邮件正文。此功能允许丰富的电子邮件内容格式。

5. **认证**：
   - 该应用程序通过`authenticate`方法处理认证。它检索发送电子邮件所需的访问令牌。

6. **错误处理**：
   - 应用程序包括基本的错误处理，如果电子邮件发送失败或收件人地址有问题，会在状态区域显示消息。

7. **退出应用程序**：
   - 关闭应用程序窗口以退出程序。

## 依赖项

- Python
- PyQt5
- Requests

在运行应用程序之前，请使用`pip`安装这些依赖项。

## 注意事项

- 使用前，请根据`config.example.ini`和`email_example.html`的示例创建自己的`config.ini`和`email.html`文件。
- 确保你拥有电子邮件服务API的正确权限和凭证。

---
