using ITworks.Brom;
using ITworks.Brom.Types;
using OneC.Interfaces;

namespace OneC.Classes;

internal class Student : IOneCActions
{
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public string Patronymic { get; set; }
    public string Gender { get; set; }
    public string TabNumber { get; set; }
    public Ссылка Link { get; set; }
    public string Type { get; set; }
    public Group Group { get; set; }

    public void AddTo1C()
    {
        БромКлиент client = Program.client;
        Запрос запрос = client.СоздатьЗапрос(@"
	    ВЫБРАТЬ
		    Студент.Ссылка КАК Ссылка
	    ИЗ
		    Справочник.Студенты КАК Студент
	    ГДЕ
		    Студент.Наименование = &наименование");
        запрос.УстановитьПараметр("наименование", $"{LastName.Trim()} {FirstName.Trim()} {Patronymic.Trim()}");
    start:
        ТаблицаЗначений результат = (ТаблицаЗначений)запрос.Выполнить();
        if (результат.Count == 0)
        {
            client.Выполнить(
                "результат = Справочники.Студенты.СоздатьЭлемент();" +
                "результат.Наименование=Параметр.Фамилия+\" \"+Параметр.Имя+\" \"+Параметр.Отчество;" +
                "результат.Имя = Параметр.Имя;" +
                "результат.Фамилия = Параметр.Фамилия;" +
                "результат.Отчество = Параметр.Отчество;" +
                "результат.ТабельныйНомер = Параметр.ТабельныйНомер;" +
                "результат.Группа=Справочники.Группы.НайтиПоНаименованию(Параметр.Группа);" +
                "результат.Пол=Параметр.Пол;" +
                "результат.Тип=Параметр.Тип;" +
                "результат.Записать();",
                new Структура("Имя,Фамилия,Отчество,Пол,Тип,Группа,ТабельныйНомер", new object[] { FirstName, LastName, Patronymic, Gender, Type, Group.Name, TabNumber }));
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