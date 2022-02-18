using System;
using System.Collections.Generic;

namespace MPL1 {
	class Program {		
		static void Main(string[] args) {
			ExcelController excelController = new ExcelController();
			ContentController contentController = new ContentController();
			WordController wordController = new WordController();

			try {
				Console.WriteLine("Считывание настроек");
				excelController.GetInfo(contentController, wordController);
				Console.WriteLine("Считывание настроек завершено");

				Console.WriteLine("Генерация открыток");
				Tuple<List<PostCard>, List<List<string>>, CelebrationsDictionary> postcardsData = contentController.CreatePostcardsContent();
				Console.WriteLine("Генерация открыток завершена");

				Console.WriteLine("Создание открыток");
				wordController.CreatePostcards(postcardsData.Item1, postcardsData.Item2, postcardsData.Item3);
				Console.WriteLine("Создание открыток завершено");
			} catch (Exception e) {
				Console.WriteLine($"Произошла ошибка {e.Message}");
			}
		}
	}
}
