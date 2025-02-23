帮我写一个电脑应用，实现如下描述的功能。

## 产品概述
- 工具名称：AI勘误器_Word版
- 功能简介：用户上传 word 文档，以`。`为分隔符把文档中的文字拆分为短句，每个短句发送给 AI API 进行错别字勘误，如果有错误则给这句话增加批注告知错误内容。
- 运行环境：工具需要被封装成可以在任意 Win 和 Mac 电脑上的可执行的程序，带界面。

## 界面和功能详述

- 工具包含一个文档上传、基础操作、过程日志显示的操作界面和一个模型信息配置界面；
- 用户可在模型配置界面填写 API 和自定义的system prompt；
- 用户可在操作界面打开本地文件里的 word 文档、开始勘误和停止操作；
- 开始勘误文档后，以`。`为分隔符，逐句读取文档的内容，标记当前读取的句子所处位置（可以考虑通过解析 XML元数）。把读取到的句子发送给大模型进行勘误，按下面逻辑来处理
  - 使用这个默认的 system Prompt(不含标签)来请求 API：
    <system_prompt>
      作为一个细致耐心的文字秘书，对下面的句子进行错别字检查，按如下结构以 JOSN 格式输出：
      {
      "content_0":"原始句子",
      "wrong":true,//是否有需要被修正的错别字，布尔类型
      "annotation":"",//批注内容，string类型。如果wrong为true给出修正的解释；如果 wrong 字段为 false，则为空值
      "content_1":""//修改后的句子，string类型。如果wrong为false则留空
      }
    </system_prompt>
  - 读取大模型 API 返回的内容，如果json 中`wrong`字段为 true 则将`annotation`添加为句子的批注；如果 wrong 字段为 false 则继续提取下一个句子。
- 完成文档所有句子勘误后，将文档保存为副本，命名为`{原始文档名}_AI勘误`，如存在同名文档则增加随机命名后缀；
- 工具界面显示勘误进度条和勘误日志，日志包括：
  <log>
  正在勘误：{被勘误的句子内容}
  没有错误/发现错误：{annotation的值}
  已在文档中批注(如有)
  </log>
- 日志可以导出为txt文档，用户选择[导出全部]或[仅错误]部分后将日志txt保存到勘误的Word文档同路径下

## 模型 API 调用和返回示例
调用示例
```
curl https://api.deepseek.com/chat/completions \
-H "Content-Type: application/json" \
-H "Authorization: Bearer %3CDeepSeek API Key%3E" \
-d '{
"model": "deepseek-chat",
"messages": [
{"role": "system", "content": "You are a helpful assistant."},
{"role": "user", "content": "Hello!"}
],
"stream": false
}'
```
其中`curl`、`model`和`API Key`的值可以在模型配置界面配置，user 角色的 Content 值为提取的句子。
API 返回示例，提取其中choices[0]message.content的值
```
{
"id": "930c60df-bf64-41c9-a88e-3ec75f81e00e",
"choices": [
{
"finish_reason": "stop",
"index": 0,
"message": {
"content": "Hello! How can I help you today?",
"role": "assistant"
}
}
],
"created": 1705651092,
"model": "deepseek-chat",
"object": "chat.completion",
"usage": {
"completion_tokens": 10,
"prompt_tokens": 16,
"total_tokens": 26
}
}
```
