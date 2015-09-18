using System;
using System.Collections.Concurrent;
using System.Collections.Generic;

namespace Translated.MateCAT.LegacyOfficeConverter.Converters
{
    public class PooledConverter<T>: IConverter where T:IConverter, new()
    {
        private BlockingCollection<T> pool;

        private bool disposed = false;


        public PooledConverter(int instancesCount)
        {
            pool = new BlockingCollection<T>();
            for (int i = 0; i < instancesCount; i++)
            {
                pool.Add(new T());
                // TODO: if supported by the instance, do the sanity
                // check. If the instance is not working stop
                // everything and notify administrators, because 
                // something very strange is happening.
            }
        }

        public bool Convert(string sourceFilePath, int sourceFormat, string targetFilePath, int targetFormat)
        {
            // Weird way to set variable to "null"
            T instance = default(T);
            try
            {
                instance = pool.Take();
                return instance.Convert(sourceFilePath, sourceFormat, targetFilePath, targetFormat);
            }
            finally
            {
                // Weird way to check that the variable is null. Sadly,
                // because of generics you can't make this cleaner.
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
        
        /// <summary>
        /// Important: this method is NOT thread-safe. 
        /// Be sure nobody is using this class before calling it.
        /// </summary>
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
