using ITworks.Brom;
using ITworks.Brom.Types;
using OneC.Interfaces;

namespace OneC.Classes;

internal class Department : IOneCActions
{
    public string Name { get; set; }
    public Ссылка Link { get; set; }
    public void AddTo1C()
    {
        БромКлиент client = Program.client;
        Запрос запрос = client.СоздатьЗапрос(@"
	    ВЫБРАТЬ
		    Подразделение.Ссылка КАК Ссылка
	    ИЗ
		    Справочник.Подразделения КАК Подразделение
	    ГДЕ
		    Подразделение.Наименование = &наименование");
        запрос.УстановитьПараметр("наименование", Name);
        start:
        ТаблицаЗначений результат = (ТаблицаЗначений)запрос.Выполнить();
        if (результат.Count == 0)
        {
            client.Выполнить(@"
                результат = Справочники.Подразделения.СоздатьЭлемент();
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