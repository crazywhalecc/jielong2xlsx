## 脚本形式使用方法

```bash
composer update
```

安装 PHPSpreadsheet 可能需要 gd、fileinfo 扩展安装。

将接龙文本写入文件 `jielong.txt` 中

```bash
php convert.php
```

如果需要统计 `酒精1瓶 防护服2套` 这样类似更多列的话，可以这样写：

```php
php convert.php 酒精 防护服
```

## 注意事项

- 支持户型：x期y号楼z室，或不带期（小区几期目前默认是根据我的小区来的）
- 附加信息中的解析关键词要跟在参数后面才能起效
