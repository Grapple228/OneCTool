using ITworks.Brom;
using ITworks.Brom.Types;
using OneC.Interfaces;

namespace OneC.Classes;

internal class Company : IOneCActions
{
    public string Name { get; set; }
    public Ссылка Link { get; set; }
    public void AddTo1C()
    {
        БромКлиент client = Program.client;
        Запрос запрос = client.СоздатьЗапрос(@"
	    ВЫБРАТЬ
		    Компания.Ссылка КАК Ссылка
	    ИЗ
		    Справочник.Компании КАК Компания
	    ГДЕ
		    Компания.Наименование = &наименование");
        запрос.УстановитьПараметр("наименование", Name);
        start:
        ТаблицаЗначений результат = (ТаблицаЗначений)запрос.Выполнить();
        if (результат.Count == 0)
        {
            client.Выполнить(@"
                результат = Справочники.Компании.СоздатьЭлемент();
                результат.Наименование=Параметр;
                результат.Записать();
            ", Name);
            goto start;
        }
        else
        {
            foreach (СтрокаТаблицыЗначений a in результат)
            {
                foreach (KeyValuePair<string, object> b in a)
                {
                    Link = (Ссылка)b.Value;
                }
            }
        }
    }
}