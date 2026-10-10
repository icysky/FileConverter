# File Converter 测试文档

这是一个 **Markdown → Word** 转换测试，支持 *斜体*、***粗斜体***、~~删除线~~ 和 `inline code`。

## 列表

- 无序列表项 1
- 无序列表项 2
    - 嵌套项目 A
    - 嵌套项目 B

1. 有序列表第一项
2. 有序列表第二项
    1. 嵌套有序项

## 代码块

```csharp
public class Hello
{
    public static void Main()
    {
        Console.WriteLine("Hello FileConverter!");
    }
}
```

## 表格

| 名称 | 类型 | 说明 |
| --- | --- | --- |
| ffmpeg | 中间件 | 音视频转换 |
| ImageMagick | 中间件 | 图片转换 |
| **Markdig** | NuGet | Markdown 解析 |

## 引用与分隔线

> 这是一段引用文字，
> 支持多行。

---

## 链接

访问 [GitHub](https://github.com/Tichau/FileConverter) 获取源码，或访问 <https://file-converter.io>。

### 结语

转换完成 ✅
