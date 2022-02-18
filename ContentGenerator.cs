using System;
using System.Collections.Generic;
using System.Linq;

namespace MPL1 {
	class CacheElement : IEquatable<CacheElement>, IComparable<CacheElement> {
		public int elementValue { get; private set; }
		public int value { get; private set; }

		public CacheElement(int elementValue, int value) {
			this.elementValue = elementValue;
			this.value = value;
		}

		public void IncrementValue() {
			value++;
		}

		public override int GetHashCode() {
			int prime = 31;
			int result = 1;
			result = prime * result + (int)(elementValue ^ ((uint)elementValue >> 32));
			result = prime * result + (int)(value ^ ((uint)value >> 32));
			return result;
		}

		public bool Equals(CacheElement phraseCache) {
			if (phraseCache == null) {
				return false;
			}

			return phraseCache.elementValue == elementValue && phraseCache.value == value;
		}

		public override bool Equals(object obj) {
			if (obj == null) {
				return false;
			}

			CacheElement phraseCache = obj as CacheElement;

			if (phraseCache == null) {
				return false;
			} else {
				return Equals(phraseCache);
			}
		}

		public int CompareTo(CacheElement phraseCache) {
			if (phraseCache == null) {
				return 1;
			}

			if (value == phraseCache.value) {
				return elementValue.CompareTo(phraseCache.elementValue);
			} else {
				return value.CompareTo(phraseCache.value);
			}
		}

		public static bool operator > (CacheElement phraseCache1, CacheElement phraseCache2) {
			return phraseCache1.CompareTo(phraseCache2) > 0;
		}

		public static bool operator < (CacheElement phraseCache1, CacheElement phraseCache2) {
			return phraseCache1.CompareTo(phraseCache2) < 0;
		}

		public static bool operator >= (CacheElement phraseCache1, CacheElement phraseCache2) {
			return phraseCache1.CompareTo(phraseCache2) >= 0;
		}

		public static bool operator <= (CacheElement phraseCache1, CacheElement phraseCache2) {
			return phraseCache1.CompareTo(phraseCache2) <= 0;
		}

		public static bool operator == (CacheElement phraseCache1, CacheElement phraseCache2) { 
			if(((object) phraseCache1) == null || ((object) phraseCache2) == null) {
				return Object.Equals(phraseCache1, phraseCache2);
			}

			return phraseCache1.Equals(phraseCache2);
		}

		public static bool operator != (CacheElement phraseCache1, CacheElement phraseCache2) {
			if (((object)phraseCache1) == null || ((object)phraseCache2) == null) {
				return !Object.Equals(phraseCache1, phraseCache2);
			}

			return !phraseCache1.Equals(phraseCache2);
		}
	}

	class FrequencyUniqueDictionary {
		private List<CacheElement> minimalGroupFrequencyCache;
		private List<List<CacheElement>> phrasesFrequencyCache;
		private Dictionary<int, int> groupToIndex;
		private Dictionary<CacheElement, Dictionary<CacheElement, Dictionary<CacheElement, bool>>> usedUniquePhrasesCache;

		public List<List<string>> phrases { get; private set; }

		public FrequencyUniqueDictionary() {
			minimalGroupFrequencyCache = new List<CacheElement>();
			phrasesFrequencyCache = new List<List<CacheElement>>();
			groupToIndex = new Dictionary<int, int>();
			phrases = new List<List<string>>();
			usedUniquePhrasesCache = new Dictionary<CacheElement, Dictionary<CacheElement, Dictionary<CacheElement, bool>>>();
		}

		public bool GroupExist(int group) {
			return groupToIndex.ContainsKey(group);
		}

		private void CreateNewGroup(int group) {
			phrases.Add(new List<string>());
			phrasesFrequencyCache.Add(new List<CacheElement>());
			
			int index = phrases.Count - 1;

			groupToIndex[group] = index;
			minimalGroupFrequencyCache.Add(new CacheElement(index, 0));
		}

		public void AddNewPhrase(int group, string text) {
			if (!GroupExist(group)) {
				CreateNewGroup(group);
			}

			int index = groupToIndex[group];
			phrases[index].Add(text);
			phrasesFrequencyCache[index].Add(new CacheElement(phrases[index].Count - 1, 0));
		}

		public int GroupsCount() {
			return minimalGroupFrequencyCache.Count;
		}

		public int GetUniqueTripletsCount() {
			int possibleUniqueCombinations = 0;
			int groups = phrases.Count;

			int[] phrasesInGroupsCount = new int[groups];

			for (int i = 0; i < groups; i++) {
				phrasesInGroupsCount[i] = phrases[i].Count;
			}

			for (int i = 0; i < groups - 2; i++) {
				for (int j = i + 1; j < groups - 1; j++) {
					for (int k = j + 1; k < groups; k++) {
						possibleUniqueCombinations += phrasesInGroupsCount[i] * phrasesInGroupsCount[j] * phrasesInGroupsCount[k];
					}
				}
			}

			return possibleUniqueCombinations;
		}

		public Tuple<int, int>[] GenerateUniquePhrase() {
			for (int i = 0; i < minimalGroupFrequencyCache.Count - 2; i++) {
				for (int j = i + 1; j < minimalGroupFrequencyCache.Count - 1; j++) {
					for (int k = j + 1; k < minimalGroupFrequencyCache.Count; k++) {
						int group1 = minimalGroupFrequencyCache[i].elementValue;
						int group2 = minimalGroupFrequencyCache[j].elementValue;
						int group3 = minimalGroupFrequencyCache[k].elementValue;

						for (int m = 0; m < phrasesFrequencyCache[group1].Count; m++) {
							for (int n = 0; n < phrasesFrequencyCache[group2].Count; n++) {
								for (int p = 0; p < phrasesFrequencyCache[group3].Count; p++) {
									int group1Element = phrasesFrequencyCache[group1][m].elementValue;
									int group2Element = phrasesFrequencyCache[group2][n].elementValue;
									int group3Element = phrasesFrequencyCache[group3][p].elementValue;

									List<CacheElement> tripletGroup = new List<CacheElement>(3) { 
										new CacheElement(group1, group1Element),
										new CacheElement(group2, group2Element),
										new CacheElement(group3, group3Element)
									};

									tripletGroup.Sort((cache1, cache2) => cache1.elementValue.CompareTo(cache2.elementValue));

									if (usedUniquePhrasesCache.ContainsKey(tripletGroup[0]) &&
										usedUniquePhrasesCache[tripletGroup[0]].ContainsKey(tripletGroup[1]) &&
										usedUniquePhrasesCache[tripletGroup[0]][tripletGroup[1]].ContainsKey(tripletGroup[2])) {
										continue;
									}

									if (!usedUniquePhrasesCache.ContainsKey(tripletGroup[0])) {
										usedUniquePhrasesCache[tripletGroup[0]] = new Dictionary<CacheElement, Dictionary<CacheElement, bool>>();
									}

									if (!usedUniquePhrasesCache[tripletGroup[0]].ContainsKey(tripletGroup[1])) {
										usedUniquePhrasesCache[tripletGroup[0]][tripletGroup[1]] = new Dictionary<CacheElement, bool>();
									}

									usedUniquePhrasesCache[tripletGroup[0]][tripletGroup[1]][tripletGroup[2]] = true;

									phrasesFrequencyCache[group1][m].IncrementValue();
									phrasesFrequencyCache[group2][n].IncrementValue();
									phrasesFrequencyCache[group3][p].IncrementValue();

									minimalGroupFrequencyCache[i].IncrementValue();
									minimalGroupFrequencyCache[j].IncrementValue();
									minimalGroupFrequencyCache[k].IncrementValue();

									phrasesFrequencyCache[group1].Sort();
									phrasesFrequencyCache[group2].Sort();
									phrasesFrequencyCache[group3].Sort();

									minimalGroupFrequencyCache.Sort();

									Tuple<int, int>[] triplets = new Tuple<int, int>[3] {
										new Tuple<int, int>(group1, group1Element),
										new Tuple<int, int>(group2, group2Element),
										new Tuple<int, int>(group3, group3Element)
									};

									return triplets;
								}
							}
						}
					}
				}
			}

			return null;
		}
	}

	class PostCard {
		public string name { get; set; }
		public string treatment { get; set; }
		public Tuple<int, int>[] triplets { get; set; }

		public PostCard(string name, string treatment) {
			this.name = name;
			this.treatment = treatment;
		}

		public void AddTriplet(Tuple<int, int>[] triplets) {
			this.triplets = triplets;
		}
	}

	class CelebrationDictionaryElement {
		public string name;
		public string templatePath;

		public CelebrationDictionaryElement(string name, string templatePath) {
			this.name = name;
			this.templatePath = templatePath;
		}
	}

	class CelebrationsDictionary {
		public Dictionary<string, CelebrationDictionaryElement> celebrations;
		public Dictionary<string, List<int>> postcardsToCelebrations;

		public CelebrationsDictionary() {
			celebrations = new Dictionary<string, CelebrationDictionaryElement>();
			postcardsToCelebrations = new Dictionary<string, List<int>>();
		}

		public void AddNewCelebration(string name, string id, string templatePath) {
			celebrations[id] = new CelebrationDictionaryElement(name, templatePath);
			postcardsToCelebrations[id] = new List<int>();
		}

		public void AddPostcardToCelebration(string celebration, int postcardIndex) {
			postcardsToCelebrations[celebration].Add(postcardIndex);
		}
	}

	class ContentController {
		public List<PostCard> postcards { get; private set; } = new List<PostCard>(0);
		public FrequencyUniqueDictionary frequencyDictionary { get; private set; } = new FrequencyUniqueDictionary();
		public CelebrationsDictionary celebrationDictionary { get; private set; } = new CelebrationsDictionary();

		public void AddNewCelebration(string name, string id, string templatePath) {
			celebrationDictionary.AddNewCelebration(name, id, templatePath);
		}

		public bool AddNewPostCard(string name, string celebration, string treatment) {
			if (name == null || celebration == null || treatment == null) {
				return false;
			}

			postcards.Add(new PostCard(name, treatment));
			AddPostcardToCelebration(celebration, postcards.Count - 1);
			return true;
		}

		public void AddPostcardToCelebration(string celebration, int postcardIndex) {
			celebrationDictionary.AddPostcardToCelebration(celebration, postcardIndex);
		}

		public void AddNewPhrase(int group, string text) {
			frequencyDictionary.AddNewPhrase(group, text);
		}

		public bool CheckUnique() {
			return postcards.Count <= frequencyDictionary.GetUniqueTripletsCount();
		}

		public bool CheckGroupsCount() {
			return frequencyDictionary.GroupsCount() >= 3;
		}


		public Tuple<List<PostCard>, List<List<string>>, CelebrationsDictionary> CreatePostcardsContent() {
			if (!CheckGroupsCount()) {
				throw new Exception("Недостаточно уникальных групп");
			}

			if (!CheckUnique()) {
				throw new Exception("Недостаточно уникальных фраз для генерации уникальных поздравлений");
			}

			for (int i = 0; i < postcards.Count; i++) {
				postcards[i].AddTriplet(frequencyDictionary.GenerateUniquePhrase());
			}

			return new Tuple<List<PostCard>, List<List<string>>, CelebrationsDictionary>(postcards, frequencyDictionary.phrases, celebrationDictionary);
		}
	}
}
