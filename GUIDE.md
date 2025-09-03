## Dify 插件开发用户指南

您好，看来您已经创建了一个插件，现在让我们开始开发吧！

### 选择您想开发的插件类型

在开始之前，您需要了解一些插件类型的基础知识。Dify 中的插件支持扩展以下能力：
- **工具**：像谷歌搜索、Stable Diffusion 等工具提供商，可用于执行特定任务。
- **模型**：像 OpenAI、Anthropic 等模型提供商，您可以使用它们的模型来增强 AI 能力。
- **端点**：类似于 Dify 中的服务 API 和 Kubernetes 中的 Ingress，您可以将一个 HTTP 服务扩展为端点，并使用自己的代码控制其逻辑。

根据您想扩展的能力，我们将插件分为三种类型：**工具**、**模型**和**扩展**。
- **工具**：它是一个工具提供商，但不仅限于工具，您还可以在其中实现一个端点。例如，如果您要构建一个 Discord 机器人，就需要 `发送消息` 和 `接收消息` 功能，此时 **工具** 和 **端点** 都是必需的。
- **模型**：仅作为模型提供商，不允许扩展其他功能。
- **扩展**：其他情况下，您可能只需要一个简单的 HTTP 服务来扩展功能，此时 **扩展** 就是您的正确选择。

我相信您在创建插件时已经选择了正确的类型，如果没有，您可以稍后通过修改 `manifest.yaml` 文件来更改。

### 清单文件

现在您可以编辑 `manifest.yaml` 文件来描述您的插件，以下是其基本结构：
- version(版本号，必填)：插件的版本。
- type(类型，必填)：插件的类型，目前仅支持 `plugin`，未来将支持 `bundle`。
- author(字符串，必填)：作者，即市场中的组织名称，也应与仓库所有者一致。
- label(标签，必填)：多语言名称。
- created_at(RFC3339 格式，必填)：创建时间，市场要求创建时间必须早于当前时间。
- icon(资源，必填)：图标路径。
- resource (对象)：要申请的资源
  - memory (int64)：最大内存使用量，主要与无服务器 SaaS 上的资源申请相关，单位为字节。
  - permission(对象)：权限申请
    - tool(对象)：反向调用工具权限
      - enabled (布尔值)
    - model(对象)：反向调用模型权限
      - enabled(布尔值)
      - llm(布尔值)
      - text_embedding(布尔值)
      - rerank(布尔值)
      - tts(布尔值)
      - speech2text(布尔值)
      - moderation(布尔值)
    - node(对象)：反向调用节点权限
      - enabled(布尔值) 
    - endpoint(对象)：允许注册端点权限
      - enabled(布尔值)
    - app(对象)：反向调用应用权限
      - enabled(布尔值)
    - storage(对象)：申请持久化存储权限
      - enabled(布尔值)
      - size(int64)：最大允许的持久化内存，单位为字节
- plugins(对象，必填)：插件扩展特定能力的 YAML 文件列表，为插件包中的绝对路径。如果您需要扩展模型，需要定义一个类似 openai.yaml 的文件，并在此处填写路径，且该路径下的文件必须存在，否则打包将失败。
  - 格式
    - tools(字符串列表)：扩展的工具供应商，详细格式请参考 [工具指南](https://docs.dify.ai/plugins/schema-definition/tool)
    - models(字符串列表)：扩展的模型供应商，详细格式请参考 [模型指南](https://docs.dify.ai/plugins/schema-definition/model)
    - endpoints(字符串列表)：扩展的端点供应商，详细格式请参考 [端点指南](https://docs.dify.ai/plugins/schema-definition/endpoint)
  - 限制
    - 不允许同时扩展工具和模型
    - 不允许没有扩展
    - 不允许同时扩展模型和端点
    - 目前每种扩展类型最多仅支持一个供应商
- meta(对象)
  - version(版本号，必填)：清单文件格式版本，初始版本为 0.0.1
  - arch(字符串列表，必填)：支持的架构，目前仅支持 amd64 和 arm64
  - runner(对象，必填)：运行时配置
    - language(字符串)：目前仅支持 Python
    - version(字符串)：语言版本，目前仅支持 3.12
    - entrypoint(字符串)：程序入口，在 Python 中应为 main

### 安装依赖

- 首先，您需要一个 Python 3.11 及以上版本的环境，因为我们的 SDK 有此要求。
- 然后，安装依赖：
    ```bash
    pip install -r requirements.txt
    ```
- 如果您想添加更多依赖，可以将它们添加到 `requirements.txt` 文件中。一旦您在 `manifest.yaml` 文件中将运行器设置为 Python，`requirements.txt` 将自动生成并用于打包和部署。

### 实现插件

现在您可以开始实现自己的插件了。通过以下示例，您可以快速了解如何实现自己的插件：
- [OpenAI](https://github.com/langgenius/dify-plugin-sdks/tree/main/python/examples/openai)：模型提供商的最佳实践
- [谷歌搜索](https://github.com/langgenius/dify-plugin-sdks/tree/main/python/examples/google)：工具提供商的简单示例
- [Neko](https://github.com/langgenius/dify-plugin-sdks/tree/main/python/examples/neko)：端点组的有趣示例

### 测试和调试插件

您可能已经注意到插件根目录下有一个 `.env.example` 文件，只需将其复制为 `.env` 并填写相应的值。如果您想在本地调试插件，需要设置一些环境变量。
- `INSTALL_METHOD`：将其设置为 `remote`，您的插件将通过网络连接到 Dify 实例。
- `REMOTE_INSTALL_URL`：Dify 实例中插件守护进程服务的调试主机和端口 URL，例如 `debug.dify.ai:5003`。[Dify SaaS](https://debug.dify.ai) 或 [自托管 Dify 实例](https://docs.dify.ai/en/getting-started/install-self-hosted/readme) 均可使用。
- `REMOTE_INSTALL_KEY`：您应该从使用的 Dify 实例中获取调试密钥。在插件管理页面的右上角，您可以看到一个带有 `debug` 图标的按钮，点击它即可获取密钥。

运行以下命令启动您的插件：
