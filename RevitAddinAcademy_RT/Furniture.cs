
using System;
using System.Collections.Generic;
using System.Linq;


namespace RevitAddinAcademy_RT
{
    public class Furniture
    {
        public string Name { get; set; }
        public string FamilyName { get; set; }
        public string FamilyType { get; set; }

        public Furniture(string name, string familyName, string familyType)
        {
            Name = name;
            FamilyName = familyName;
            FamilyType = familyType;
        }
        
        //Select a Furniture object from a list of all furniture objects by name
        public static Furniture GetFurnitureByName(List<Furniture> allFurniture, string name)
        {
            foreach(Furniture furniture in allFurniture)
            {
                if (furniture.Name == name)
                    return furniture;
            }
            return null;
        }
    }

    public class FurnitureSet
    {
        public string Set { get; set; }
        public string RoomType { get; set; }
        public List<Furniture> FurnitureList { get; set; }
        public double FurnitureCount { get; set; }


        public FurnitureSet(string set, string roomType, string furnitureList, List<Furniture> allFurniture)
        {
            Set = set;
            RoomType = roomType;
            FurnitureList = GetFurnitureList(allFurniture, furnitureList);
            FurnitureCount = Convert.ToDouble(FurnitureList.Count);
        }

        //Create a list of Furniture objects selected from all available Furniture objects
        //Inputs - allFurniture: List of all furniture objects; furnString: comma separated list of furniture in the furniture set
        private List<Furniture> GetFurnitureList(List<Furniture> allFurniture, string furnString)
        {
            List<string> strings = furnString.Split(',').ToList();
            List<string> names = new List<string>();
            foreach(string s in strings)
            {
                string name = s.Trim();
                if (name != null && name != "")
                    names.Add(name);
            }

            List<Furniture> result = new List<Furniture>();

            foreach(string name in names)
            {
                Furniture curFurn = Furniture.GetFurnitureByName(allFurniture, name);

                result.Add(curFurn);
            }
            return result;
        }

        //Select a Furniture set object by name
        public static FurnitureSet GetFurnitureSetByName(List<FurnitureSet> allSets, string name)
        {
            foreach (FurnitureSet set in allSets)
            {
                if (set.Set == name)
                    return set;
            }
            return null;
        }
    }

}
