using System;
using System.Collections;
using System.Text;
using System.IO;

namespace RPAUserAdminOJT.Controllers.ActiveDirectory.DnParser
{
    public class RDNComponent
    {
        #region Enumerations

        public enum RDNValueType
        {
            StringValue,
            HexValue
        };

        #endregion

        #region Data Members

        private string componentType;
        private string componentValue;
        private RDNValueType componentValueType;
        private int hashCode;

        #endregion

        #region Properties

        public string ComponentType
        {
            get
            {
                return componentType;
            }
        }

        public string ComponentValue
        {
            get
            {
                return componentValue;
            }
        }

        public RDNValueType ComponentValueType
        {
            get
            {
                return componentValueType;
            }
        }

        #endregion

        #region Constructors

        internal RDNComponent(string ComponentType, string ComponentValue, RDNValueType ComponentValueType)
        {
            componentType = ComponentType;
            componentValue = ComponentValue;
            componentValueType = ComponentValueType;

            GenerateHashCode();
        }

        #endregion

        #region Methods

        public override bool Equals(object obj)
        {
            if (object.ReferenceEquals(obj, null))
            {
                return false;
            }

            if (this.hashCode != obj.GetHashCode())
                return false;

            if (obj is RDNComponent)
            {
                RDNComponent rdnObj = (RDNComponent)obj;

                return (rdnObj.componentValueType == this.componentValueType &&
                    string.Compare(rdnObj.componentType, this.componentType, true) == 0 &&
                    string.Compare(rdnObj.componentValue, this.componentValue, true) == 0
                    );
            }
            else
            {
                return false;
            }
        }

        private void GenerateHashCode()
        {
            hashCode = 0x48012e7a;

            hashCode ^= this.componentValueType.GetHashCode();

            if (this.componentType != null)
                hashCode ^= this.componentType.ToLower().GetHashCode();

            if (this.componentValue != null)
                hashCode ^= this.componentValue.ToLower().GetHashCode();
        }

        public override int GetHashCode()
        {
            return hashCode;
        }

        public override string ToString()
        {
            return ToString(DN.DefaultEscapeChars);
        }

        public string ToString(EscapeChars escapeChars)
        {
            if (this.componentValueType == RDNValueType.HexValue)
            {
                return this.componentType + "=" + this.componentValue;
            }
            else
            {
                return this.componentType + "=" + EscapeValue(this.componentValue, escapeChars);
            }
        }


        #endregion

        #region Static Methods

        public static string EscapeValue(string s, EscapeChars escapeChars)
        {
            StringBuilder ReturnValue = new StringBuilder();

            for (int i = 0; i < s.Length; i++)
            {
                if (RDN.IsSpecialChar(s[i]) || ((i == 0 || i == s.Length - 1) && s[i] == ' '))
                {
                    if ((escapeChars & EscapeChars.SpecialChars) != EscapeChars.None)
                        ReturnValue.Append('\\');

                    ReturnValue.Append(s[i]);
                }
                else if (s[i] < 32 && ((escapeChars & EscapeChars.ControlChars) != EscapeChars.None))
                {
                    ReturnValue.AppendFormat("\\{0:X2}", (int)s[i]);
                }
                else if (s[i] >= 128 && ((escapeChars & EscapeChars.MultibyteChars) != EscapeChars.None))
                {
                    byte[] Bytes = Encoding.UTF8.GetBytes(new char[] { s[i] });

                    foreach (byte b in Bytes)
                    {
                        ReturnValue.AppendFormat("\\{0:X2}", b);
                    }
                }
                else
                {
                    ReturnValue.Append(s[i]);
                }
            }

            return ReturnValue.ToString();
        }


        #endregion

        #region Overloaded Operators

        public static bool operator ==(RDNComponent obj1, RDNComponent obj2)
        {
            if (object.ReferenceEquals(obj1, null))
            {
                return (object.ReferenceEquals(obj2, null));
            }
            else
            {
                return obj1.Equals(obj2);
            }
        }

        public static bool operator !=(RDNComponent obj1, RDNComponent obj2)
        {
            return (!(obj1 == obj2));
        }

        #endregion
    }

    public class RDNComponentList : IEnumerable
    {
        #region Data Members

        private RDNComponent[] components;

        #endregion

        #region Constructors

        internal RDNComponentList(RDNComponent[] components)
        {
            this.components = components;
        }

        #endregion

        #region Properties

        public int Length
        {
            get
            {
                return components.Length;
            }
        }

        public RDNComponent this[int index]
        {
            get
            {
                return components[index];
            }
        }

        #endregion

        #region IEnumerable Members

        public IEnumerator GetEnumerator()
        {
            return components.GetEnumerator();
        }

        #endregion
    }
}