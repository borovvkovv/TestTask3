using Castle.MicroKernel.Registration;
using Castle.Windsor;
using Spectre.Console;
using System.Data;
using System.Text;
using Castle.MicroKernel;
using TestTask3;
using Path = System.IO.Path;
using Table = Spectre.Console.Table;

class Program
{
    delegate void OptionFunction();

    static void Main(string[] args)
    {
        Stack<OptionFunction> CallStack = new Stack<OptionFunction>();
        IWindsorContainer windsorContainer = new WindsorContainer();
        Console.OutputEncoding = Encoding.UTF8;

        void GetFile()
        {
            string excelPath = string.Empty;
            var jsonWriter = new JsonWriter("appsettings.json");
            var config = new Config(jsonWriter);
            var excelPathExists = false;
            var checkConfig = true;

            while (!excelPathExists)
            {
                var isExcelPathFromConfig = config.TryGetParam("excelPath", out excelPath);
                if (checkConfig && isExcelPathFromConfig)
                {
                    if (!AnsiConsole.Confirm($"Продолжить работу с файлом \"{excelPath}\"?", true))
                        excelPath = AnsiConsole.Ask<string>("Введите полный путь до файла:");
                }
                else
                {
                    excelPath = AnsiConsole.Ask<string>("Введите полный путь до файла:");
                }

                excelPath = excelPath.Trim().Replace("\"", string.Empty);
                if (!Path.Exists(excelPath))
                {
                    AnsiConsole.WriteLine("Файл не найден!");
                    checkConfig = false;
                }
                else
                {
                    config.SetConfigParam("excelPath", excelPath);
                    excelPathExists = true;
                }
            }

            windsorContainer.Register(
                Component.For<ITableLoader>().ImplementedBy<OpenXmlLoader>().Named(Guid.NewGuid().ToString()).IsDefault().DependsOn(Dependency.OnValue("path", excelPath))
            );

            InvokeOptionFunction(GetMainMenu);
        }

        void GetMainMenu()
        {
            var result = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Выберите команду:")
                    .AddChoices("Поиск клиентов по товару",
                        "Изменение контактного лица клиента",
                        "Определить золотого клиента",
                        "[red]Назад[/]"));

            if (result == "Поиск клиентов по товару")
            {
                InvokeOptionFunction(ShowClientsByProduct);
            }
            else if (result == "Изменение контактного лица клиента")
            {
                InvokeOptionFunction(AlterClientContactInfo);
            }
            else if (result == "Определить золотого клиента")
            {
                InvokeOptionFunction(GetGoldenClient);
            }
            else if (result == "[red]Назад[/]")
            {
                Back();
            }
        }

        void ShowClientsByProduct()
        {
            DataTable products = null;
            DataTable clients = null;
            DataTable reqests = null;
            var openXml = windsorContainer.Resolve<ITableLoader>();
            try
            {
                products = openXml.LoadTable("Товары");
            }
            catch (Exception e)
            {
                AnsiConsole.MarkupLine("[red]Не удалось загрузить таблицу Товары.[/]");
                AnsiConsole.MarkupLine($"[red]{e.Message}[/]");
                return;
            }

            var selectedProductName = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Выберите товар:")
                    .AddChoices(products.AsEnumerable().Select(row => row.Field<string>("Наименование"))));

            try
            {
                clients = openXml.LoadTable("Клиенты");
                reqests = openXml.LoadTable("Заявки");
            }
            catch (Exception e)
            {
                AnsiConsole.MarkupLine("[red]Не удалось загрузить таблицу.[/]");
                AnsiConsole.MarkupLine($"[red]{e.Message}[/]");
                return;
            }

            var resultTable = new Table();
            resultTable.AddColumn("Контактное лицо (ФИО)");
            resultTable.AddColumn("Код клиента");
            resultTable.AddColumn("Количество товара");
            resultTable.AddColumn("Цена");
            resultTable.AddColumn("Дата заказа");

            var selectedProduct = products.Select($"[Наименование] = '{selectedProductName}'").FirstOrDefault();
            var productCode = selectedProduct.Field<string>("Код товара");
            var productPrice = selectedProduct.Field<string>("Цена товара за единицу");
            var productRequests = reqests.Select($"[Код товара] = '{productCode}'");
            foreach (var request in productRequests)
            {
                var clientCode = request.Field<string>("Код клиента");
                var clientInfo = clients.Select($"[Код клиента] = '{clientCode}'").FirstOrDefault();
                var clientFio = clientInfo.Field<string>("Контактное лицо (ФИО)");
                var orderNumber = request.Field<string>("Требуемое количество");
                var orderDate = request.Field<string>("Дата размещения");

                resultTable.AddRow(clientFio, clientCode, orderNumber, productPrice, orderDate);
            }

            AnsiConsole.Write(resultTable);
        }

        void AlterClientContactInfo()
        {
            DataTable clients = null;
            var openXml = windsorContainer.Resolve<ITableLoader>();
            try
            {
                clients = openXml.LoadTable("Клиенты");
            }
            catch (Exception e)
            {
                AnsiConsole.MarkupLine("[red]Не удалось загрузить таблицу Клиенты.[/]");
                AnsiConsole.MarkupLine($"[red]{e.Message}[/]");
                return;
            }

            var selectedClientName = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Выберите контактное лицо:")
                    .AddChoices(clients.AsEnumerable().Select(row => row.Field<string>("Контактное лицо (ФИО)"))));

            var clientInfoTable = new Table();
            clientInfoTable.AddColumn("Код клиента");
            clientInfoTable.AddColumn("Наименование организации");
            clientInfoTable.AddColumn("Адрес");
            clientInfoTable.AddColumn("Контактное лицо (ФИО)");
            var selectedClient = clients.Select($"[Контактное лицо (ФИО)] = '{selectedClientName}'").FirstOrDefault();
            var clientCode = selectedClient.Field<string>("Код клиента");
            var clientOrgname = selectedClient.Field<string>("Наименование организации");
            var clientAddress = selectedClient.Field<string>("Адрес");
            clientInfoTable.AddRow(clientCode, clientOrgname, clientAddress, selectedClientName);
            AnsiConsole.Write(clientInfoTable);

            var newClientOrgName = AnsiConsole.Prompt(
                new TextPrompt<string>("Введите новое наименование организации. Оставьте пустую строку, если изменение не требуется.")
                    .DefaultValue(clientOrgname)
                    .ShowDefaultValue());
            selectedClient.SetField("Наименование организации", newClientOrgName);

            var newClientName = AnsiConsole.Prompt(
                new TextPrompt<string>("Введите новое ФИО. Оставьте пустую строку, если изменение не требуется.")
                    .DefaultValue(selectedClientName)
                    .ShowDefaultValue());
            selectedClient.SetField("Контактное лицо (ФИО)", newClientName);

            if (newClientOrgName != clientOrgname || newClientName != selectedClientName)
            {
                try
                {
                    openXml.SaveTable(clients, "Клиенты");
                    clients = openXml.LoadTable("Клиенты");
                }
                catch (Exception e)
                {
                    AnsiConsole.WriteException(e);
                    return;
                }

                clientInfoTable.RemoveRow(0);
                selectedClient = clients.Select($"[Код клиента] = '{clientCode}'").FirstOrDefault();
                clientOrgname = selectedClient.Field<string>("Наименование организации");
                var clientName = selectedClient.Field<string>("Контактное лицо (ФИО)");
                clientInfoTable.AddRow(clientCode, clientOrgname, clientAddress, clientName);
                AnsiConsole.Write(clientInfoTable);
            }
            else
                AnsiConsole.MarkupLine("[red]Данные из изменились.[/]");
        }

        void GetGoldenClient()
        {
            var result = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Пожалуйста выберите за какой период найти золотого клиента:")
                    .AddChoices("За месяц",
                        "За год",
                        "[red]Назад[/]"));

            if (result == "За месяц")
            {
                InvokeOptionFunction(GetGoldenClientOfMonth);
            }
            else if (result == "За год")
            {
                InvokeOptionFunction(GetGoldenClientOfYear);
            }
            else if (result == "[red]Назад[/]")
            {
                Back();
            }
        }

        void GetGoldenClientOfMonth()
        {
            var months = new List<string>();
            var minMonth = new DateTime(DateTime.Now.Year, 1, 1).AddYears(-3);
            var maxMonth = new DateTime(DateTime.Now.Year, 1, 1);
            var currentMonth = minMonth;
            while (currentMonth < maxMonth)
            {
                months.Add(currentMonth.ToString(("MM.yyyy")));
                currentMonth = currentMonth.AddMonths(1);
            }
            var selectedMonth = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Выберите месяц:")
                    .AddChoices(months));

            try
            {
                var table = GetGoldenClientOfPeriod(selectedMonth);
                if (table.Rows.Count > 0)
                    AnsiConsole.Write(table);
                else
                    AnsiConsole.MarkupLine("[red]За указанный год не было заказов.[/]");

            }
            catch (Exception e)
            {
                AnsiConsole.MarkupLine($"[red]{e.Message}[/]");
            }
        }

        void GetGoldenClientOfYear()
        {
            var years = new List<string>();
            var minYear = DateTime.Now.Year - 3;
            var maxYear = DateTime.Now.Year;
            var currentYear = minYear;
            while (currentYear < maxYear)
            {
                years.Add(currentYear.ToString());
                currentYear++;
            }
            var selectedYear = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Выберите год:")
                    .AddChoices(years));

            try
            {
                var table = GetGoldenClientOfPeriod(selectedYear);
                if (table.Rows.Count > 0)
                    AnsiConsole.Write(table);
                else
                    AnsiConsole.MarkupLine("[red]За указанный год не было заказов.[/]");

            }
            catch (Exception e)
            {
                AnsiConsole.MarkupLine($"[red]{e.Message}[/]");
            }
        }

        Table GetGoldenClientOfPeriod(string selectedPeriod)
        {
            DataTable orders = null;
            DataTable clients = null;
            try
            {
                var openXml = windsorContainer.Resolve<ITableLoader>();
                orders = openXml.LoadTable("Заявки");
                clients = openXml.LoadTable("Клиенты");
            }
            catch (Exception e)
            {
                throw new Exception($"Не удалось загрузить таблицу. {e.Message}");
            }

            var clientInfoTable = new Table();
            var ordersInSelectedMonth = orders.Select($"[Дата размещения] Like '%{selectedPeriod}'");
            var ordersPerClient = ordersInSelectedMonth.GroupBy(x => x.Field<string>("Код клиента")).OrderByDescending(x => x.Count()).FirstOrDefault();
            if (ordersPerClient != null)
            {
                clientInfoTable.AddColumns(["Код клиента", "Наименование организации", "Адрес", "Контактное лицо (ФИО)", "Количество заказов"]);
                var clientCode = ordersPerClient.Key;
                var client = clients.Select($"[Код клиента] = '{clientCode}'").FirstOrDefault();
                var clientOrgname = client.Field<string>("Наименование организации");
                var clientAddress = client.Field<string>("Адрес");
                var clientName = client.Field<string>("Контактное лицо (ФИО)");
                var ordersCount = ordersPerClient.Select(x => x.Field<string>("Требуемое количество")).Count().ToString();
                clientInfoTable.AddRow(clientCode, clientOrgname, clientAddress, clientName, ordersCount);
            }

            return clientInfoTable;
        }

        void InvokeOptionFunction(OptionFunction foo)
        {
            CallStack.Push(foo);
            foo();

            if (AnsiConsole.Confirm("Назад?", true))
                Back();
            else
                Repeat();
        }

        void Back()
        {
            CallStack.Pop();
            var lastFunction = CallStack.Pop();
            InvokeOptionFunction(lastFunction);
        }

        void Repeat()
        {
            var lastFunction = CallStack.Pop();
            InvokeOptionFunction(lastFunction);
        }

        InvokeOptionFunction(GetFile);
    }
    
}