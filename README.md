SharepointLibrary
=================
Sample of use:
<pre><code>
const string СсылкаПортала = @"http://xxxxx/_layouts/viewlsts.aspx";
public void ДанныеСпискаДоговоров()
{
    try
    {
        var портал = new SharepointAccess(СсылкаПортала);
        var список = портал.ПолучитьСписки().FirstOrDefault(x => x.Title.Contains("Docs"));
        var данные = портал.ПолучитьДанныеСписка(список);
        if (данные != null) Debug.WriteLine("Количество строк = " + данные.Count);
    }
    catch (Exception ex) {  }
}
</code></pre>
