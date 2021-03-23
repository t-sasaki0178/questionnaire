# 適当な文字数でテキストを返すやつ。なんか動いてないっぽいので後でなんとかしたい
def truncate(string,length,ellipsis='...'):
    return string[:length] + (ellipsis if string[length:] else '')