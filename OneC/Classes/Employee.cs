using ITworks.Brom;
using ITworks.Brom.Types;
using OneC.Interfaces;

namespace OneC.Classes;

internal class Employee : IOneCActions
{
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public string Patronymic { get; set; }
    public string Gender { get; set; }
    public string TabNumber { get; set; }
    public Company Company { get; set; }
    public Department Department { get; set; }
    public Post Post { get; set; }
    public Ссылка Link { get; set; }
    public void AddTo1C()
    {
        БромКлиент client = Program.client;
        Запрос запрос = client.СоздатьЗапрос(@"
	    ВЫБРАТЬ
		    Сотрудник.Ссылка КАК Ссылка
	    ИЗ
		    Справочник.СотрудникиКолледжа КАК Сотрудник
	    ГДЕ
		    Сотрудник.ТабельныйНомер = &ТабельныйНомер");
        запрос.УстановитьПараметр("ТабельныйНомер", $"{TabNumber}");
        start:
        ТаблицаЗначений результат = (ТаблицаЗначений)запрос.Выполнить();
        if (результат.Count == 0)
        {
            client.Выполнить(
                "результат = Справочники.СотрудникиКолледжа.СоздатьЭлемент();" +
                "результат.Наименование=Параметр.Фамилия+\" \"+Параметр.Имя+\" \"+Параметр.Отчество;" +
                "результат.Имя = Параметр.Имя;" +
                "результат.Фамилия = Параметр.Фамилия;" +
                "результат.Отчество = Параметр.Отчество;" +
                "результат.ТабельныйНомер = Параметр.ТабельныйНомер;" +
                "результат.Компания=Справочники.Компании.НайтиПоНаименованию(Параметр.Компания);" +
                "результат.Подразделение=Справочники.Подразделения.НайтиПоНаименованию(Параметр.Подразделение);" +
                "результат.Должность=Справочники.Должность.НайтиПоНаименованию(Параметр.Должность);" +
                "результат.Пол=Параметр.Пол;" +
                "результат.Записать();",
                new Структура("Имя,Фамилия,Отчество,Пол,Компания,Подразделение,Должность,ТабельныйНомер", new object[] { FirstName, LastName, Patronymic, Gender, Company.Name, Department.Name, Post.Name,TabNumber }));
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