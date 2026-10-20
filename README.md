# coc7-xlsx-web

浏览器内运行 COC7 车卡 xlsx（丛雨的版本），解决在没有 Excel 的环境使用的问题。使用完全一致的公式和验证规则，导出时以空白卡原始文件为基础修改，支持 CY23.2 和 CY23Final 两个版本。目前只在 Safari 环境里测试过。

## Usage

用浏览器打开 `dist/COC7车卡.html` 即可。
- 然后像 Excel 一样修改，存储在浏览器本地。
- "列表" 管理多张卡，新建时可以选 CY23.2 或 CY23Final。
- "打开 xlsx"可以导入空白卡或者别人在 Excel 里改过的卡，"导出 xlsx" 生成文件。
- "设置"里可以转换版本、撤销工作表保护、打开紧凑 UI。

## Build

```sh
npm install
npm run build
npm test
```

## About Template

内置的 `COC7空白卡CY23.2.xlsx` 和 `COC7空白卡CY23Final.xlsx` 是丛雨制作的（原卡由秋叶EXODUS制作），版权归原作者，不在本仓库代码的授权范围内。


## License

代码采用 [MIT](LICENSE) 许可。
