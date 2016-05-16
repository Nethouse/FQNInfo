using System;
using System.ComponentModel;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using EnvDTE;
using VSLangProj;
using IExtenderProvider = EnvDTE.IExtenderProvider;

namespace fqninfo
{
    [ComVisible(true)]
    public interface IFQNProperty
    {
        string FullyQualifiedName { get; }
    }
    
    [ComVisible(true)]
    public class FQNProperty : IFQNProperty, IDisposable
    {
        private readonly IExtenderSite _extenderSite;
        private readonly int _cookie;
        private bool _disposed = false;

        [DisplayName("Fully Qualified Name")]
        [Category("Misc")]
        [Description("The Fully Qualified Name (FQN) of the assembly.")]
        public string FullyQualifiedName { get; }


        public FQNProperty(string fileName, IExtenderSite extenderSite, int cookie)
        {
            if (fileName == null)
            {
                throw new ArgumentNullException("fileName");
            }

            if (extenderSite == null)
            {
                throw new ArgumentNullException("extenderSite");
            }

            _extenderSite = extenderSite;
            _cookie = cookie;

            try
            {
                var fqn = AssemblyName.GetAssemblyName(fileName);
                FullyQualifiedName = fqn.ToString();
            }
            catch
            {
                FullyQualifiedName = "<Could not resolve FQN>";
            }
        }

        ~FQNProperty()
        {
            Dispose();
        }


        public void Dispose()
        {
            Dispose(true);
            // take the instance off of the finalization queue.
            GC.SuppressFinalize(this);
        }

        private void Dispose(bool disposing)
        {
            if (_disposed) return;
            if (disposing && _cookie != 0)
            {
                _extenderSite.NotifyDelete(_cookie);
            }
            _disposed = true;
        }
    }


    public class FQNPropertyExtender : IExtenderProvider
    {
        public const string SupportedExtenderName = "FQNPropertyExtender";
        public const string SupportedExtenderCATID = PrjBrowseObjectCATID.prjCATIDCSharpReferenceBrowseObject;
        private IFQNProperty _extender;

        public object GetExtender(string ExtenderCATID, string ExtenderName, object ExtendeeObject, IExtenderSite ExtenderSite,
            int Cookie)
        {
            if (CanExtend(ExtenderCATID, ExtenderName, ExtendeeObject))
            {
                var obj = (Reference)ExtendeeObject;
                _extender = new FQNProperty(obj.Path, ExtenderSite, Cookie);
                return _extender;
            }
            return null;
        }

        public bool CanExtend(string ExtenderCATID, string ExtenderName, object ExtendeeObject)
        {
            return ExtenderName == SupportedExtenderName // check if the correct extener is requested
                   && string.Equals(ExtenderCATID, SupportedExtenderCATID, StringComparison.OrdinalIgnoreCase) // check if the correct CATID is requested
                   && string.Equals(ExtendeeObject.GetPropertyValue<string>("ExtenderCATID"), SupportedExtenderCATID, StringComparison.OrdinalIgnoreCase); // check whether the extended object really has this CATID
        }
    }

    public static class TypeDescriptorSupport
    {
        public static T GetPropertyValue<T>(this object source, string propertyName)
        {
            if (source == null)
            {
                throw new ArgumentNullException("source");
            }

            if (propertyName == null)
            {
                throw new ArgumentNullException("propertyName");
            }

            object value = null;

            PropertyDescriptor property = TypeDescriptor.GetProperties(source)[propertyName];
            if (property != null)
            {
                value = property.GetValue(source);
            }

            return value != null ? (T)Convert.ChangeType(value, typeof(T), CultureInfo.InvariantCulture) : default(T);
        }
    }
}
