using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Runtime.Serialization.Json;

namespace DataArrange.Storages
{
    public class Storage
    {
        [Serializable]
        public struct DataItem
        {
            public string Key;
            public string Value;
            public DataItem newkey(string key)
            {
                Value = key; return this;
            }
        }
        [Serializable]
        public struct DataArea
        {
            public string User;
            public List<DataItem> Items;
        }
        [Serializable]
        public struct DataAreas
        {
            public string Owner;
            public List<DataArea> Areas;
        }
        public DataAreas data;
        public bool FirstStore = false;
        public List<DataItem> GetUserData(string user)
        {
            int u = getuser(user);
            return (data.Areas[u].Items);
        }
        public void RemoveUser(string user)
        {
            int u = getuser(user);
            data.Areas.RemoveAt(u);
        }
        public void Remove(string user, string key)
        {
            int u = getuser(user);
            int i = data.Areas[u].Items.FindIndex(m => m.Key == key);
            if (i == -1) { return; }
            data.Areas[u].Items.RemoveAt(i);
        }
        public void putkey(string user, int index, string value, bool autostore = true)
        {
            int u = getuser(user);
            data.Areas[u].Items[index] = data.Areas[u].Items[index].newkey(value);
            if (autostore) { Store(); }
        }
        public void putkey(string user, string key, string value, bool autostore = true)
        {
            int u = getuser(user);
            int i = data.Areas[u].Items.FindIndex(m => m.Key == key);
            if (i == -1) { data.Areas[u].Items.Add(new DataItem { Key = key, Value = value }); if (autostore) { Store(); } return; }
            data.Areas[u].Items[i] = data.Areas[u].Items[i].newkey(value);
            if (autostore) { Store(); }
        }
        public string getkey(string user, int index)
        {
            int u = getuser(user);
            return (data.Areas[u].Items[index].Value);
        }
        public string getkey(string user, string key)
        {
            int u = getuser(user);
            return (data.Areas[u].Items.Find(m => m.Key == key).Value);
        }
        public int getuser(string user)
        {
            int u = data.Areas.FindIndex(m => m.User == user);
            if (u == -1)
            {
                DataArea da = new DataArea();
                da.User = user; da.Items = new List<DataItem>();
                data.Areas.Add(da);
                u = data.Areas.FindIndex(m => m.User == user);
            }
            return u;
        }
        public Storage(string storageID)
        {
            if (Directory.Exists(@"C:\DataArrange\") == false) { Directory.CreateDirectory(@"C:\DataArrange\"); }
            data = new DataAreas { Owner = storageID };
            data.Areas = new List<DataArea>();
            Restore();
        }
        public void Store()
        {
            DataContractJsonSerializer w = new DataContractJsonSerializer(typeof(DataAreas));
            FileStream f = File.Create(@"C:\DataArrange\" + data.Owner + "-userdata.json");
            w.WriteObject(f, data);
            f.Close();
        }
        public void Restore()
        {
            if (!File.Exists(@"C:\DataArrange\" + data.Owner + "-userdata.json")) { FirstStore = true; return; }
            DataContractJsonSerializer r = new DataContractJsonSerializer(typeof(DataAreas));
            FileStream f = File.Open(@"C:\DataArrange\" + data.Owner + "-userdata.json", FileMode.Open);
            data = (DataAreas)r.ReadObject(f);
            f.Close();
        }

    }
}
