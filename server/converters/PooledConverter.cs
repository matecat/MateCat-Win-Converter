using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LegacyOfficeConverter
{
    class PooledConverter<T>: IConverter where T:IConverter, new()
    {
        private BlockingCollection<T> pool;

        private bool disposed = false;


        public PooledConverter(int instancesCount)
        {
            pool = new BlockingCollection<T>();
            for (int i = 0; i < instancesCount; i++)
            {
                pool.Add(new T());
            }
        }

        public string Convert(string path)
        {
            T instance = default(T);
            try
            {
                instance = pool.Take();
                string convertedFilePath = instance.Convert(path);
                return convertedFilePath;
            }
            finally
            {
                if (!EqualityComparer<T>.Default.Equals(instance, default(T)))
                {
                    pool.Add(instance);
                }
            }
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
                while (pool.TryTake(out instance))
                {
                    if (instance is IDisposable)
                        ((IDisposable) instance).Dispose();
                }
                pool.Dispose();
            }

            disposed = true;
        }

        ~PooledConverter()
        {
            Dispose(false);
        }
    }
}
