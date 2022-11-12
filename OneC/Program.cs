using ITworks.Brom;
using ITworks.Brom.Types;
using ITworks.Brom.Metadata;
using ITworks.Brom.SOAP;
using System.Xml.Linq;


namespace Program;

static class Program
{
	public static БромКлиент client;
    public static List<Student> students;
    public static List<Employee> employees;
    public static List<Group> groups;
    public static List<Company> companies;
    public static List<Post> posts;
    public static List<Department> departments;
    public static void Main(string[] args)
	{
		const string PATH = @"C:\Users\Green Apple\Desktop\Data\";

        // Подключение к 1с
        client = new БромКлиент(@"
        	Публикация		= http://localhost/Buh/;
        	Пользователь	= bromuser;
        	Пароль			= bromuser
        ");

        students = new();
        employees = new();
        groups = new();
        companies = new();
        departments = new();
        posts = new();

        // Заполнение студентов
        string[] studentsFromFile = File.ReadAllLines($"{PATH}Студенты.csv");
        for (int i = 1; i < studentsFromFile.Length; i++)
        {
	        string[] studentData = studentsFromFile[i].Split(';');
        	string groupString = studentData[5].Replace("Студент ", "").Trim();
        	groupString = groupString == "" ? "Не указано" : groupString;
        	Group? group = groups.SingleOrDefault(x => x.Name == groupString);
        	if (group == null)
        	{
        		group = new Group(groupString);
        		groups.Add(group);
        	}
        	Student student = new Student()
        	{
        		FirstName = studentData[0].Trim(),
        		LastName = studentData[1].Trim(),
        		Patronymic = studentData[2].Trim(),
        		Type = studentData[4].Trim(),
        		Group = group,
                TabNumber = studentData[3].Trim() == "" ? "Не указано" : studentData[3].Trim(),
        		Gender = studentData[8].Trim(),
        	};
        	students.Add(student);
        }

        // Заполнение сотрудников
        string[] employeesFromFile = File.ReadAllLines($"{PATH}Сотрудники.csv");
        for (int i = 1; i < employeesFromFile.Length; i++)
        {
            string[] employeeData = employeesFromFile[i].Split(';');
            Company? company = companies.SingleOrDefault(x => x.Name == employeeData[6].Trim());
            if(company == null)
            {
                company = new Company()
                {
                    Name = employeeData[6].Trim() == "" ? "Не указано" : employeeData[6].Trim(),
                };
                companies.Add(company);
            }
            Post? post = posts.SingleOrDefault(x => x.Name == employeeData[5].Trim());
            if (post == null)
            {
                post = new Post()
                {
                    Name = employeeData[5].Trim() == "" ? "Не указано" : employeeData[5].Trim(),
                };
                posts.Add(post);
            }
            Department? department = departments.SingleOrDefault(x => x.Name == employeeData[4].Trim());
            if (department == null)
            {
                department = new Department()
                {
                    Name = employeeData[4].Trim() == "" ? "Не указано" : employeeData[4].Trim(),
                };
                departments.Add(department);
            }

            Employee employee = new()
            {
                FirstName= employeeData[0].Trim(),
                LastName= employeeData[2].Trim(),
                Patronymic= employeeData[1].Trim(),
                Gender= employeeData[8].Trim(),
                Company= company,
                Department= department,
                Post = post,
                TabNumber= employeeData[3].Trim() == "" ? "Не указано" : employeeData[3].Trim(),                
            };
            employees.Add(employee);
        }

        // Добавление в 1С
        foreach (var group in groups)
        {
        	group.AddTo1C();
        }
        foreach (var student in students)
        {
            student.AddTo1C();
        }
        foreach (var post in posts)
        {
            post.AddTo1C();
        }
        foreach (var company in companies)
        {
            company.AddTo1C();
        }
        foreach (var department in departments)
        {
            department.AddTo1C();
        }
        foreach (var employee in employees)
        {
            employee.AddTo1C();
        }
    }
}

interface IOneCActions
{
    public Ссылка Link { get; set; }
    public void AddTo1C();
}

class Post : IOneCActions
{
    public string Name { get; set; }
    public Ссылка Link { get; set; }
    public void AddTo1C()
    {
        БромКлиент client = Program.client;
        Запрос запрос = client.СоздатьЗапрос(@"
	    ВЫБРАТЬ
		    ДолжностьУ.Ссылка КАК Ссылка
	    ИЗ
		    Справочник.Должность КАК ДолжностьУ
	    ГДЕ
		    ДолжностьУ.Наименование = &наименование");
        запрос.УстановитьПараметр("наименование", Name);
        start:
        ТаблицаЗначений результат = (ТаблицаЗначений)запрос.Выполнить();
        if (результат.Count == 0)
        {
            client.Выполнить(@"
                результат = Справочники.Должность.СоздатьЭлемент();
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
class Department : IOneCActions
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
class Company : IOneCActions
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

class Employee : IOneCActions
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

class Student : IOneCActions
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

class Group : IOneCActions
{
    public string Name { get; set; }
    public Ссылка Link { get; set; }

    public Group(string name)
    {
        Name = name;
    }

    public void AddTo1C()
    {
        БромКлиент client = Program.client;
        Запрос запрос = client.СоздатьЗапрос(@"
	    ВЫБРАТЬ
		    Группа.Ссылка КАК Ссылка
	    ИЗ
		    Справочник.Группы КАК Группа
	    ГДЕ
		    Группа.Наименование = &наименование");
        запрос.УстановитьПараметр("наименование", Name);
    start:
        ТаблицаЗначений результат = (ТаблицаЗначений)запрос.Выполнить();
        if (результат.Count == 0)
        {
            client.Выполнить(@"
                результат = Справочники.Группы.СоздатьЭлемент();
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
