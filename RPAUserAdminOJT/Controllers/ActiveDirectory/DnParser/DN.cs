using System;
using System.Collections;
using System.Text;

namespace RPAUserAdminOJT.Controllers.ActiveDirectory.DnParser
{
    public class DN
    {
        #region Enumerations

        private enum ParserState { LookingForSeparator, InQuotedString };

        #endregion

        #region Static Data Members

        private static EscapeChars defaultEscapeChars = EscapeChars.ControlChars | EscapeChars.SpecialChars;

        #endregion

        #region Data Members

        private RDNList rDNs;
        private EscapeChars escapeChars;
        private int hashCode;

        #endregion

        #region Properties

        public EscapeChars CharsToEscape
        {
            get
            {
                return escapeChars;
            }
            set
            {
                escapeChars = value;
            }
        }

        public RDNList RDNs
        {
            get
            {
                return rDNs;
            }
        }

        public DN Parent
        {
            get
            {
                if (rDNs.Length >= 1)
                {
                    RDN[] parentRDNs = new RDN[rDNs.Length - 1];

                    for (int i = 0; i < parentRDNs.Length; i++)
                    {
                        parentRDNs[i] = rDNs[i + 1];
                    }

                    return new DN(new RDNList(parentRDNs), this.CharsToEscape);
                }
                else
                {
                    throw new InvalidOperationException("Can't get the parent of an empty DN");
                }
            }
        }

        #endregion

        #region Constructors

        public DN(string dnString) : this(dnString, DefaultEscapeChars) { }

        public DN(string dnString, EscapeChars escapeChars)
        {
            if (dnString == null)
            {
                throw new ArgumentNullException("dnString");
            }

            this.escapeChars = escapeChars;

            ParseDN(dnString);

            GenerateHashCode();
        }

        private DN(RDNList rdnList, EscapeChars escapeChars)
        {
            this.escapeChars = escapeChars;

            rDNs = rdnList;

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

            if (hashCode != obj.GetHashCode())
                return false;

            if (obj is DN)
            {
                DN dnObj = (DN)obj;

                if (dnObj.rDNs.Length == this.rDNs.Length)
                {
                    for (int i = 0; i < this.rDNs.Length; i++)
                    {
                        if (!(dnObj.rDNs[i].Equals(this.rDNs[i])))
                            return false;
                    }
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        public override int GetHashCode()
        {
            return hashCode;
        }

        private void GenerateHashCode()
        {
            hashCode = 0x28f527b4;

            for (int i = 0; i < this.rDNs.Length; i++)
            {
                hashCode ^= this.rDNs[i].GetHashCode();
            }
        }

        public override string ToString()
        {
            return ToString(this.escapeChars);
        }

        public string ToString(EscapeChars escapeChars)
        {
            StringBuilder ReturnValue = new StringBuilder();

            foreach (RDN rdn in RDNs)
            {
                ReturnValue.Append(rdn.ToString(escapeChars));
                ReturnValue.Append(",");
            }

            if (ReturnValue.Length > 0)
                ReturnValue.Length--;

            return ReturnValue.ToString();
        }


        private void ParseDN(string dnString)
        {
            if (dnString.Length == 0)
            {
                rDNs = new RDNList(new RDN[0]);
                return;
            }

            ArrayList rawRDNs = new ArrayList();
            ParserState state = ParserState.LookingForSeparator;
            StringBuilder rawRDN = new StringBuilder();


            for (int position = 0; position < dnString.Length; ++position)
            {
                switch (state)
                {
                    #region case ParserState.LookingForSeparator:
                    case ParserState.LookingForSeparator:
                        if (dnString[position] == ',' || dnString[position] == ';')
                        {
                            rawRDNs.Add(rawRDN.ToString());
                            rawRDN.Length = 0;
                        }
                        else
                        {
                            rawRDN.Append(dnString[position]);

                            if (dnString[position] == '\\')
                            {
                                try
                                {
                                    rawRDN.Append(dnString[++position]);
                                }
                                catch (IndexOutOfRangeException)
                                {
                                    throw new ArgumentException("Invalid DN: DNs aren't allowed to end with an escape character.", dnString);
                                }

                            }
                            else if (dnString[position] == '"')
                            {
                                state = ParserState.InQuotedString;
                            }
                        }

                        break;

                    #endregion

                    #region case ParserState.InQuotedString:

                    case ParserState.InQuotedString:
                        rawRDN.Append(dnString[position]);

                        if (dnString[position] == '\\')
                        {
                            try
                            {
                                rawRDN.Append(dnString[++position]);
                            }
                            catch (IndexOutOfRangeException)
                            {
                                throw new ArgumentException("Invalid DN: DNs aren't allowed to end with an escape character.", dnString);
                            }
                        }
                        else if (dnString[position] == '"')
                            state = ParserState.LookingForSeparator;
                        break;

                        #endregion
                }
            }

            rawRDNs.Add(rawRDN.ToString());

            if (state == ParserState.InQuotedString)
                throw new ArgumentException("Invalid DN: Unterminated quoted string.", dnString);

            RDN[] results = new RDN[rawRDNs.Count];

            for (int i = 0; i < results.Length; i++)
            {
                results[i] = new RDN(rawRDNs[i].ToString());
            }

            rDNs = new RDNList(results);
        }

        public bool Contains(DN childDN)
        {
            if (childDN.rDNs.Length > this.rDNs.Length)
            {
                int Offset = childDN.rDNs.Length - this.rDNs.Length;

                for (int i = 0; i < this.rDNs.Length; i++)
                {
                    if (childDN.rDNs[i + Offset] != this.rDNs[i])
                        return false;
                }

                return true;
            }
            else
            {
                return false;
            }
        }

        public DN GetChild(string childRDN)
        {
            DN childDN = new DN(childRDN);

            if (childDN.rDNs.Length > 0)
            {
                RDN[] fullPath = new RDN[this.RDNs.Length + childDN.rDNs.Length];

                for (int i = 0; i < childDN.rDNs.Length; i++)
                {
                    fullPath[i] = childDN.rDNs[i];
                }
                for (int j = 0; j < this.rDNs.Length; j++)
                {
                    fullPath[j + childDN.rDNs.Length] = this.rDNs[j];
                }

                return new DN(new RDNList(fullPath), this.CharsToEscape);
            }
            else
            {
                return this;
            }
        }


        #endregion

        #region Overloaded Operators

        public static bool operator ==(DN obj1, DN obj2)
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

        public static bool operator !=(DN obj1, DN obj2)
        {
            return (!(obj1 == obj2));
        }

        #endregion

        #region Static Properties

        public static EscapeChars DefaultEscapeChars
        {
            get
            {
                return defaultEscapeChars;
            }
            set
            {
                defaultEscapeChars = value;
            }
        }

        #endregion
    }

    [Flags]
    public enum EscapeChars
    {
        None = 0,
        ControlChars = 1,
        SpecialChars = 2,
        MultibyteChars = 4
    }
}
