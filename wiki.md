## EUCなど
- CSV出力するためにクリックしたい要素
```html
<SPAN id=export_b_item title="" class=action><INPUT onclick="setupEvent(event);if(!checkAndResetBlurCheckResult()){return false;}event_onClick_withIFrame('export_b','export_bAction','',0)" onkeyup="setupEvent(event);cursor_onKeyUp('EXPORT_B')" onfocus="setupEvent(event);event_onFocus('action_export_b')" id=sidEXPORT_B class=IoAction onkeydown="setupEvent(event);return event_onKeyDown4Item('EXPORT_B', '')" type=button value=ＩＤリスト出力 name=action_export_b> 
</SPAN>
```

## IEモードでHTMLの要素を取得したいとき
```
"C:\Windows\System32\F12\IEChooser.exe" 
```