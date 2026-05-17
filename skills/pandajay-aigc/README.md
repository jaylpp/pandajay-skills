# pandajay-aigc - 降低文本 AIGC 检测率

使用 Paper Panda API 对文本进行智能改写，支持知网、维普、朱雀等 11 种改写类型。

## 首次使用：配置 API Key

### 1. 获取 API Key

访问 **https://paperpanda.cn** 注册账号，然后在「API密钥管理」中创建一个新的 API Key。

### 2. 配置 API Key

```bash
python3 scripts/rewrite.py --set-key <你的API Key>
```

配置完成后即可正常使用。

---

## 使用示例

```bash
# 通用降AI
python3 scripts/rewrite.py "要改写的文本" 3 ai-reduce

# 朱雀降AI
python3 scripts/rewrite.py "要改写的文本" 11 ai-reduce

# 知网双降（降AI+降重）
python3 scripts/rewrite.py "要改写的文本" 7 ai-duplicate
```

## 改写类型

| ID | 名称 | 适用场景 |
|----|------|----------|
| 1 | 知网降AI | 知网平台AI检测 |
| 2 | 维普降AI | 维普平台AI检测 |
| 3 | 通用降AI | 默认选项，通用场景 |
| 5 | 通用降重 | 专注降重服务 |
| 6 | 格子达降AI | 格子达平台AI检测 |
| 7 | 双降知网 | 知网平台同时降AI+降重 |
| 8 | 双降维普 | 维普平台同时降AI+降重 |
| 9 | 双降通用 | 通用场景同时降AI+降重 |
| 10 | 双降格子达 | 格子达平台同时降AI+降重 |
| 11 | 朱雀降AI | 朱雀平台AI检测 |
| 12 | 特价降AI | 追求性价比 |