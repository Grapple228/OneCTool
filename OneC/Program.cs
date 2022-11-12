using ITworks.Brom;
using OneC.Classes;

namespace OneC;

internal static class Program
{
	public static БромКлиент client;
    private static List<Student> students;
    private static List<Employee> employees;
    private static List<Group> groups;
    private static List<Company> companies;
    private static List<Post> posts;
    private static List<Department> departments;
    public static void Main(string[] args)
	{
        // Расположение плоских файлов
		const string path = @"C:\Users\Green Apple\Desktop\Data\";

        // Подключение к 1с через IIS
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
        string[] studentsFromFile = File.ReadAllLines($"{path}Студенты.csv");
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
        string[] employeesFromFile = File.ReadAllLines($"{path}Сотрудники.csv");
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
