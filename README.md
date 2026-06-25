# OutlookMail-OAuth-SMTP-ProxyMailServer

这是一个专为解决微软（Outlook / Office 365）禁用基本身份验证（Basic Authentication）而设计的本地 OAuth2 邮件代理服务器。

它可以作为一个本地网关，允许那些不支持 OAuth2 认证协议的旧版邮件客户端、打印机、扫描仪、报警系统或遗留系统，继续通过标准的用户名和密码方式登录，而在后台自动转换并使用微软最新的 OAuth2 流程安全地收发邮件。
# 📌 为什么需要它？

微软自 2022 年起已全面停用 Exchange Online / Outlook.com 的基本身份验证（密码直接登录）。所有的 SMTP、IMAP、POP3 连接都必须通过现代化身份验证（Modern Authentication / OAuth2）。

这意味着：

- 旧版系统（如一些企业内部的 ERP、遗留的 C#/.NET 4.0/Java 程序）

- 硬件设备（如多功能打印机的 Scan-to-Email 功能）

- 不支持现代化认证的客户端

现在都无法直接发送邮件。

该项目通过在本地（或局域网服务器上）搭建一个代理服务：

```
[ 传统客户端/设备 ] --(发送传统SMTP邮件)--> [ mail_api.py (本地SMTP服务器) ]
                                                   |
                                            (调用微软 Graph API)
                                                   v
[ get_refresh_token.py (负责OAuth2授权) ] --> [ 微软云端服务 (Office 365) ]
```

# ✨ 主要特性

- 协议网关：在本地提供标准的 SMTP（默认端口 1025）。

- OAuth2 令牌管理：自动处理微软 OAuth2 授权码流程，负责 Refresh Token 的本地持久化存储和自动刷新（Access Token）。

- 无缝对接：旧版客户端仅需将邮件服务器地址改为本地 IP，即可无需修改代码重构。

- 安全性：本地通信可绑定 127.0.0.1 仅允许本机访问，亦可配置局域网 TLS 加密。

# 🚀 快速开始

## 准备工作：在 Azure 注册应用程序

为了代表你的 Outlook 账户发送/接收邮件，你需要在微软 Azure 门户中注册一个应用（免费）。

登录 Azure 门户 (Azure Portal)。

导航到 Azure Active Directory（或 Microsoft Entra ID） -> 应用注册 -> 新注册。

填写应用信息：

名称：例如 Outlook-OAuth-Proxy。

受支持的账户类型：选择“任何组织目录中的账户和个人 Microsoft 账户（例如 Skype、Xbox）”（根据你的邮箱类型选择）。

重定向 URI：平台选择 Web 或 移动和桌面应用程序（建议选择公共客户端/原生），地址填写：http://localhost:8080/callback（需与配置文件中保持一致）。

创建后，记录下 应用程序(客户端) ID (Client ID)。

针对需要客户端机密（Client Secret）的流，可在 证书和密码 中生成一个新密码，并记录。

在 API 权限 中，添加以下 Microsoft Graph / Office 365 Exchange Online 权限（委托的权限）：

- offline_access (必须，用于获取刷新令牌)

- SMTP.Send (用于发送邮件)

- IMAP.AccessAsUser.All (如果需要接收邮件)

- POP.AccessAsUser.All (可选)

## 📦 安装与配置

步骤 1： 克隆仓库与安装依赖
```
git clone [https://github.com/your-username/OutlookMail-OAuth-SMTP-ProxyMailServer.git](https://github.com/your-username/OutlookMail-OAuth-SMTP-ProxyMailServer.git)
cd OutlookMail-OAuth-SMTP-ProxyMailServer
pip install -r requirements.txt
```

步骤 2： 编辑配置文件

编辑 config.txt：
```
[microsoft]
client_id = 
redirect_uri = http://localhost:8000

[tokens]
refresh_token = 
access_token =
expires_at = 
```

注意：local_password 是你为旧版邮件客户端设置的本地连接密码。客户端连接本地代理时使用此密码进行校验，验证成功后，代理服务器会使用 OAuth2 令牌与微软服务器通信。

🔑 步骤 3：初始化 OAuth 授权

首次运行需要进行一次交互式授权以获取 Refresh Token：

python get_refresh_token.py


程序会输出一个微软登录 URL。将该 URL 复制到浏览器打开，登录你的 Outlook 邮箱并同意授权。授权成功后，页面会重定向到 http://localhost:8080/callback，本地程序捕获到授权码后将自动保存 Token 至本地加密的 config.txt 中。

🏃 步骤 4：运行代理服务器

运行以下命令正式启动本地代理服务：

python mail_api.py


终端将显示服务已成功绑定本地端口：

SMTP 代理已在 127.0.0.1:1025 启动

# ⚙️ 客户端配置指南

现在你可以修改你原本的旧版客户端/打印机/系统设置：

|配置项|SMTP 服务器|SMTP 端口|IMAP 服务器|IMAP 端口|SSL/TLS 选项|用户名/邮箱|密码|
|-|-|-|-|-|-|-|-|
|原配置 (不可用)|smtp.office365.com|587|outlook.office365.com|993|STARTTLS / SSL|your-email@outlook.com|你的微软账号密码|
|新配置 (通过代理)|127.0.0.1 (或代理服务器的局域网 IP)|1025|127.0.0.1|1143|无/None (本地代理通常无需加密，若跨机器可配置本地自签名证书)|your-email@outlook.com|local_secure_password_for_client (在 config.txt 里设置的本地密码)|


# ❓ 常见问题 (FAQ)

Q: Token 会过期吗？需要每隔几天就重新授权吗？
A: 不需要。由于我们在 Azure 权限中申请了 offline_access，程序会获取到一个 Refresh Token（刷新令牌）。每次 Access Token 过期时，代理服务器会在后台自动使用 Refresh Token 换取新的 Access Token，无需人工干预。

Q: 能够部署在公网或局域网服务器上吗？
A: 可以。如果你需要为局域网内的多台物理设备（如打印机）提供服务，请将 bind_address 修改为 0.0.0.0。强烈建议在这种情况下开启 TLS 支持，并在配置文件中配置 ssl_cert 与 ssl_key 以确保局域网内的密码及邮件安全。

Q: 是否支持多账户？
A: 支持。你可以在 accounts 数组中添加多个账号。代理服务器会根据旧版客户端登录时输入的邮箱账号（Username），自动匹配并使用对应的 OAuth 令牌去连接微软服务器。

# 📄 开源许可证

本项目采用 MIT License 许可协议。
