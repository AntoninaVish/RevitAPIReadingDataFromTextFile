using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace RevitAPIReadingDataFromTextFile
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand

    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            //указываем из какого имено файла записываем данные для этого воспользуемся диалоговым окном
            //для того чтобы выбрать файл создаем переменную openFileDialog
            OpenFileDialog openFileDialog = new OpenFileDialog();

           openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); //Directory по умолчанию будет рабочий стол

            openFileDialog.Filter = "All files(*.*)|*.*"; //выделяем фильтр, хотим чтобы отображались все файлы 

            string filePath = string.Empty; //создаем переменную, которая сохранит путь к файлу 
            if (openFileDialog.ShowDialog() == DialogResult.OK) //если путь к сохранению файла указан
            {
                //тогда мы забираем этот путь из переменной openFileDialog, записаный путь записывается в свойство FileName
                filePath = openFileDialog.FileName;
            }
            //Если строка у нас null или пустая, тогда просто возвращаем и заканчиваем выполнение нашей программы
            if (string.IsNullOrEmpty(filePath))
                return Result.Cancelled;

           var lines = File.ReadAllLines(filePath).ToList(); //в любом другом случае мы забираем все строки из текстового файла

            List<RoomData> roomDataList = new List<RoomData>(); //создаем переменную списка RoomData

            //пройдемся в списке по каждой строке, создаем еще один список с данными и разделяем каждую строку по значению разделителя
            foreach (var line in lines)
            {
                List<string> values = line.Split(';').ToList();

                //добавляем данные в список
                roomDataList.Add(new RoomData
                {
                    Name = values[0],
                    Number = values[1]
                });
            }

            string roomsInfo = string.Empty;

             var rooms = new FilteredElementCollector(doc)
                .OfCategory (BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            //создаем using для того чтобы создать переменную, создавать транзакцию и записывать данные в модель
            using (var ts = new Transaction(doc, "Set parameter"))
            {
                ts.Start();//начало транзакции

                //проходимся по списку RoomData -это список данных из текстового файла
                foreach (RoomData roomData in roomDataList)
                {
                    //находим помещение с тем номером которое указано в RoomData, обращаемся ко всем помещениям
                    //и находим то помещение которое совпадает по номеру с данным из RoomData
                    Room room = rooms.FirstOrDefault(r => r.Number.Equals(roomData.Number));

                    
                    if (room == null) //если помещение не найдено, нужного номера нету
                        continue; //переходим к следующему RoomData

                    //обращаемся к параметру "Имени", устанавливаем новое значение Set(roomData.Name) из RoomData
                    room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(roomData.Name);
                }

                ts.Commit();//конец транзакции

            }



            return Result.Succeeded;
        }
    }
}
