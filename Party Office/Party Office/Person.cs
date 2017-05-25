using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Party_Office
{
    class Person
    {
        string name;
        List<int> confer;
        List<bool> preOrNot; // presentation -> 1, participant -> 0

        public Person(string Name)
        {
            name = Name;
            confer = new List<int>();
            preOrNot = new List<bool>();
        }

        public string Name { get => name; set => name = value; }
        public List<int> Confer { get => confer; set => confer = value; }
        public List<bool> PreOrNot { get => preOrNot; set => preOrNot = value; }

        public void AddConf(int ConfNo, bool PreOrNot)
        {
            confer.Add(ConfNo);
            preOrNot.Add(PreOrNot);
        }
    }
}
