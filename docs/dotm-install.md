# WordEquationNumbering.dotm 安装说明

## 生成模板

本仓库已经提供 VBA 源码：

```text
vba\EquationNumbering.bas
```

自动生成 `.dotm` 需要 Word 允许程序访问 VBA 工程。打开 Word：

```text
文件 > 选项 > 信任中心 > 信任中心设置 > 宏设置 > 信任对 VBA 项目对象模型的访问
```

然后运行：

```powershell
cd C:\Software\AI\Word公式编辑器
powershell -ExecutionPolicy Bypass -File tools\Build-Dotm.ps1
```

输出文件：

```text
release\WordEquationNumbering.dotm
```

## 手动导入源码

如果不想打开“信任对 VBA 项目对象模型的访问”，可以手动创建模板：

1. 打开 Word，新建空白文档。
2. 按 `Alt+F11` 打开 VBA 编辑器。
3. 菜单选择 `文件 > 导入文件`。
4. 选择 `vba\EquationNumbering.bas`。
5. 回到 Word，另存为 `Word 启用宏的模板 (*.dotm)`。

## 加载模板

把 `.dotm` 放到 Word 启动目录：

```text
%APPDATA%\Microsoft\Word\STARTUP
```

重启 Word 后，在 `开发工具 > 宏` 中运行这些宏：

- `InsertEquationLinePlain`
- `InsertEquationLineChapterHyphen`
- `InsertEquationLineChapterDot`
- `InsertEquationLineChapterColon`
- `InsertEquationReference`
- `RefreshEquationFields`

## 签名

这台电脑有 Office 自带的 SelfCert：

```text
C:\Program Files\Microsoft Office\root\Office16\SELFCERT.EXE
```

可以用它创建自签名证书，然后在 VBA 编辑器中：

```text
工具 > 数字签名
```

选择证书后保存模板。自签名证书只适合本机或受控环境使用。
