using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace MPL1 {
	class WordController {
		private string templatePath;		
		private string wordFont;

		public bool IsTemplateExist(string path) {
			FileInfo templateFile = new FileInfo(path);

			if (!templateFile.Exists) {
				throw new Exception($"Файла шаблона не существует");
				return false;
			}

			SetPath(path);

			return true;
		}

		public void SetPath(string templatePath) {
			this.templatePath = templatePath;
		}

		public void SetFont(string wordFont) {
			this.wordFont = wordFont;
		}


		public void CreatePostcards(List<PostCard> postcardsData, List<List<string>> phrases, CelebrationsDictionary celebrationsDictionary) {
			object missing = System.Reflection.Missing.Value;

			string date = DateTime.UtcNow.ToString("MM_dd_yyyy HH_mm_ss");

			DirectoryInfo resultDir = new DirectoryInfo($@"{Directory.GetCurrentDirectory()}\{date}");
			if (!resultDir.Exists) {
				resultDir.Create();
			}

			foreach (string key in celebrationsDictionary.celebrations.Keys) {
				FileInfo templateFile = new FileInfo($@"{Directory.GetCurrentDirectory()}\{celebrationsDictionary.celebrations[key].templatePath}");
				if (!templateFile.Exists) {
					continue;
				}

				Word.Application templateWordApp = new Word.Application();
				Word.Document templateDoc = templateWordApp.Documents.Open($@"{Directory.GetCurrentDirectory()}\{celebrationsDictionary.celebrations[key].templatePath}");

				Word.Application outputWordApp = new Word.Application();
				Word.Document outputDoc = templateWordApp.Documents.Add(ref missing, ref missing, ref missing, ref missing);

				templateDoc.Content.Copy();

				List<int> postcardIndexesList = celebrationsDictionary.postcardsToCelebrations[key];
				string celebrationName = celebrationsDictionary.celebrations[key].name;

				for (int i = 0; i < postcardIndexesList.Count; i++) {
					outputDoc.Range(outputDoc.Content.End - 1, outputDoc.Content.End - 1).Paste();

					PostCard postcard = postcardsData[postcardIndexesList[i]];
					Tuple<int, int> triplet1 = postcard.triplets[0];
					Tuple<int, int> triplet2 = postcard.triplets[1];
					Tuple<int, int> triplet3 = postcard.triplets[2];

					outputDoc.Bookmarks.get_Item("Treatment").Range.Text = postcard.treatment;
					outputDoc.Bookmarks.get_Item("Name").Range.Text = postcard.name;
					outputDoc.Bookmarks.get_Item("Celebration").Range.Text = celebrationName;
					outputDoc.Bookmarks.get_Item("Triplets").Range.Text = 
						$"{phrases[triplet1.Item1][triplet1.Item2]}, {phrases[triplet2.Item1][triplet2.Item2]}, {phrases[triplet3.Item1][triplet3.Item2]}";
				}

				outputDoc.SaveAs2($@"{resultDir.FullName}\{celebrationName}.docx");

				templateDoc.Close();
				outputDoc.Close();

				templateWordApp.Quit();
				outputWordApp.Quit();
			}
		}
	}
}
