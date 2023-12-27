using System;
using System.Linq;

namespace ExcelTask
{
    internal class Program
    {
        const string WrongChoice = "Неверная команда. Введите цифру запроса";
        static void Main(string[] args)
        {
            while (true)
            {
                ExecuteCommand(GetChoice());
                Console.WriteLine();
            }
        }

        static int GetChoice()
        {
            Console.WriteLine("Просьба указать, какое действие хотите выполнить. Введите цифру запроса");
            Console.WriteLine("1. Загрузить файл с данными");
            Console.WriteLine("2. Получить информацию о заказах по товару");
            Console.WriteLine("3. Изменить контактное лицо клиента");
            Console.WriteLine("4. Определить золотого клиента");
            Console.WriteLine("5. Закрыть программу\n");
            while (true)
            {
                var choice = Console.ReadLine();
                if (int.TryParse(choice, out var number)
                    && Enumerable.Range(1, 5).Contains(number))
                    return number;

                Console.WriteLine(WrongChoice);
            }
        }

        static void ExecuteCommand(int choice)
        {
            switch (choice)
            {
                case 1:
                    Console.WriteLine("Укажите полный путь до файла для загрузки");
                    if (Executor.Instance.LoadEntities(Console.ReadLine()))
                        Console.WriteLine("Данные успешно загружены");
                    break;
                case 2:
                    Console.WriteLine("Введите наименование товара");
                    var wareName = Console.ReadLine();
                    var orders = Executor.Instance.GetClientsByWare(wareName);
                    Console.WriteLine("Заказы по товару:");
                    foreach ((DateTime Date, Client Client, int Quantity, decimal Price) order in orders)
                    {
                        Console.WriteLine($"{order.Date.ToShortDateString()}: заказ на количество {order.Quantity} по цене {order.Price}, " +
                            $"клиент {order.Client.Name} - контактное лицо {order.Client.Contact}");
                    }
                    break;
                case 3:
                    Executor.Instance.SetNewContact(null, null);
                    break;
                case 4:
                    var dateValue = 0;
                    var isYear = true;
                    while (true)
                    {
                        Console.WriteLine("Определить за год или месяц?");
                        Console.WriteLine("1. За год");
                        Console.WriteLine("2. За месяц");
                        Console.WriteLine("3. Вернуться к запросам");
                        var choiceForGoldSting = Console.ReadLine();
                        if (int.TryParse(choiceForGoldSting, out var number)
                            && Enumerable.Range(1, 3).Contains(number))
                        {
                            if (number == 3)
                                break;

                            Console.WriteLine($"Введите {(number == 2 ? "месяц через цифру" : "год")}");
                            choiceForGoldSting = Console.ReadLine();
                            if (int.TryParse(choiceForGoldSting, out var choiceForGold))
                            {
                                dateValue = choiceForGold;
                                isYear = number == 1;
                                break;
                            }

                        }
                        Console.WriteLine(WrongChoice);
                    }
                    var goldClient = Executor.Instance.GetGoldClient(isYear, dateValue);
                    if (goldClient != null)
                        Console.WriteLine($"{Executor.Instance.GetGoldClient(isYear, dateValue)?.Name}");
                    break;
                case 5:
                    Environment.Exit(0);
                    break;
                default:
                    Console.WriteLine("Неизвестная команда");
                    break;
            }

            if (Executor.Instance.InnerException != null)
            {
                Console.WriteLine(Executor.Instance.InnerException.Message);
                Executor.Instance.InnerException = null;
            }
        }
    }
}
