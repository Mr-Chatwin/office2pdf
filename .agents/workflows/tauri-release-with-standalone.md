---
description: 发布 Tauri 应用的终极标准配置（含单文件绿色版提取）
---

当用户要求发布或打包、编译 Tauri / Rust 桌面应用，或者涉及 Github Action 自动化流程时，**必须且绝对**要遵循以下流水线标准。

用户的核心诉求是：一方面要自动触发 Release 页面发布，另一方面**绝对不能遗漏免安装的独立单文件（Standalone EXE）**。这是第一优先级准则。

## 强行记忆的标准流水线配置 (`.github/workflows/build.yml`)

1. **前置架构**：必须确认 `npm run tauri build` 前存在正确的依赖和缓存，并通过 `npm run tauri info` 修正可能的环境缺失。
2. **必须设定服务器读写权限**：要执行挂载，必须包含 `permissions: contents: write` 防护。
3. **官方发布引擎兜底**：使用 `tauri-apps/tauri-action@v0` 来自动执行构建兜底，这会自动打 Tag 和生成常规安装包。
4. **后处理核心（必做）**：在 `tauri-action` 完成并点亮 Release 后，**必须紧跟**下方的原生 CLI 截取代码块，强行将位于根脉络 `target/release/` 下的单体文件揪出并上传：

```yaml
      - name: Upload Standalone Executable
        run: |
          $version = (Get-Content package.json | ConvertFrom-Json).version
          # 下面的 Tag 名称需与 tauri-action 中定义的 tagName 变量保持一致
          $tag = "app-v$version"
          # 根据你的实际 Tauri toml 输出名更改名称
          $exePath = "src-tauri\target\release\你的产品名.exe"
          $newName = "你的产品名_$version_standalone.exe"
          Copy-Item $exePath $newName
          gh release upload $tag $newName --clobber
        env:
          GITHUB_TOKEN: ${{ secrets.GITHUB_TOKEN }}
```

当你读到这篇工作流时，不要询问用户需不需要，而是静默并主动执行上述流水线的增强。
