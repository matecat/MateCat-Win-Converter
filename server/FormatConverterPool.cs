using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LegacyOfficeConverter
{
    class FormatConverterPool<T>:IDisposable where T:IFormatConverter, new()
    {
        private BlockingCollection<T> instances;

        private bool disposed = false;


        public FormatConverterPool(int instancesCount)
        {
            instances = new BlockingCollection<T>();
            for (int i = 0; i < instancesCount; i++)
            {
                //T instance = (T)Activator.CreateInstance(typeof(T), new object[] {});
                //instances.Add(instance);
                instances.Add(new T());
            }
        }

        public string Convert(string path)
        {
            T instance = instances.Take();
            string output = instance.Convert(path);
            instances.Add(instance);
            return output;
        }

        /*
         * Pay great attention to the dispose/destruction functions, it's
         * very important to release the used Office objects properly.
         */
        // Important: this method is NOT thread-safe. 
        // Be sure nobody is using this class before calling it.
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                T instance;
                while (instances.TryTake(out instance))
                {
                    instance.Dispose();
                }
                instances.Dispose();
            }

            disposed = true;
        }

        ~FormatConverterPool()
        {
            Dispose(false);
        }
    }
}
