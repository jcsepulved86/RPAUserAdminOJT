using System;
using System.Collections;
using System.Text;
using System.IO;

namespace RPAUserAdminOJT.Controllers.ActiveDirectory.DnParser
{
    public class RDN
    {
        #region Enumerations

        private enum ParserState { DetermineAttributeType, GetTypeByOID, GetTypeByName, DetermineValueType, GetQuotedValue, GetUnquotedValue, GetHexValue };

        #endregion

        #region Data Members

        private RDNComponentList components;
        private char[] charArray = new char[1];
        private int hashCode;

        #endregion

        #region Properties


        public RDNComponentList Components
        {
            get
            {
                return components;
            }
        }

        #endregion

        #region Construtors

        internal RDN(string rdnString)
        {
            ParseRDN(rdnString);

            GenerateHashCode();
        }

        private RDN(RDNComponentList components)
        {
            this.components = components;

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

            if (obj is RDN)
            {
                RDN rdnObj = (RDN)obj;

                if (rdnObj.components.Length == this.components.Length)
                {
                    for (int i = 0; i < rdnObj.components.Length; i++)
                    {
                        if (!(rdnObj.components[i].Equals(this.components[i])))
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
            hashCode = 0x74f8149a;

            for (int i = 0; i < this.components.Length; i++)
            {
                hashCode ^= this.components[i].GetHashCode();
            }
        }

        public override string ToString()
        {
            return ToString(DN.DefaultEscapeChars);
        }

        public string ToString(EscapeChars escapeChars)
        {
            StringBuilder ReturnValue = new StringBuilder();

            foreach (RDNComponent component in components)
            {
                ReturnValue.Append(component.ToString(escapeChars));
                ReturnValue.Append("+");
            }

            if (ReturnValue.Length > 0)
                ReturnValue.Length--;

            return ReturnValue.ToString();
        }


        private void WriteUTF8Bytes(Stream s, char c)
        {
            charArray[0] = c;

            byte[] utf8Bytes = Encoding.UTF8.GetBytes(charArray);

            s.Write(utf8Bytes, 0, utf8Bytes.Length);
        }


        private void ParseRDN(string rdnString)
        {
            ArrayList rawTypes = new ArrayList();
            ArrayList rawValues = new ArrayList();
            ArrayList rawValueTypes = new ArrayList();

            MemoryStream rawData = new MemoryStream();

            ParserState state = ParserState.DetermineAttributeType;

            int position = 0;

            while (position < rdnString.Length)
            {
                switch (state)
                {
                    #region case ParserState.DetermineAttributeType:

                    case ParserState.DetermineAttributeType:

                        #region Ignore any spaces at the beginning

                        try
                        {
                            while (rdnString[position] == ' ')
                                ++position;
                        }
                        catch (IndexOutOfRangeException)
                        {
                            throw new ArgumentException("Invalid RDN: It's entirely blank spaces", rdnString);
                        }

                        #endregion

                        #region Are we looking at a letter?

                        if (IsAlpha(rdnString[position]))
                        {
                            string StartPoint = rdnString.Substring(position);

                            if (StartPoint.StartsWith("OID.") ||
                                StartPoint.StartsWith("oid."))
                            {
                                position += 4;
                                state = ParserState.GetTypeByOID;
                            }
                            else
                            {
                                state = ParserState.GetTypeByName;
                            }
                        }

                        #endregion

                        #region If not, are we looking at a digit?

                        else if (IsDigit(rdnString[position]))
                        {
                            state = ParserState.GetTypeByOID;
                        }

                        #endregion

                        #region If not, we're looking at something that shouldn't be there.

                        else
                        {
                            throw new ArgumentException("Invalid RDN: Invalid character in attribute name", rdnString);
                        }

                        #endregion

                        break;

                    #endregion

                    #region case ParserState.GetTypeByName:

                    case ParserState.GetTypeByName:

                        try
                        {
                            #region The first character needs to be a letter

                            if (IsAlpha(rdnString[position]))
                            {
                                WriteUTF8Bytes(rawData, rdnString[position]);
                                position++;
                            }
                            else
                            {
                                throw new ArgumentException("Invalid Attribute Name: Name must start with a letter", rdnString);
                            }

                            #endregion

                            #region The remaining characters can be letters, digits, or hyphens

                            while (IsAlpha(rdnString[position]) || IsDigit(rdnString[position]) || rdnString[position] == '-')
                            {
                                WriteUTF8Bytes(rawData, rdnString[position]);
                                position++;
                            }

                            #endregion

                            #region The name can be followed by any number of blank spaces

                            while (rdnString[position] == ' ')
                            {
                                position++;
                            }

                            #endregion

                            #region And it needs to end with an equals sign

                            if (rdnString[position] == '=')
                            {
                                position++;
                                rawTypes.Add(Encoding.UTF8.GetString(rawData.ToArray()));
                                rawData.SetLength(0);

                                if (position >= rdnString.Length)
                                {
                                    rawValues.Add("");
                                    rawValueTypes.Add(RDNComponent.RDNValueType.StringValue);
                                    state = ParserState.GetUnquotedValue;
                                }
                                else
                                {
                                    state = ParserState.DetermineValueType;
                                }
                            }
                            else
                            {
                                throw new ArgumentException("Invalid RDN: unterminated attribute name", rdnString);
                            }

                            #endregion
                        }
                        catch (IndexOutOfRangeException)
                        {
                            throw new ArgumentException("Invalid RDN: unterminated attribute name", rdnString);
                        }

                        break;

                    #endregion

                    #region case ParserState.GetTypeByOID:

                    case ParserState.GetTypeByOID:

                        try
                        {
                            #region The first character needs to be a digit

                            if (IsDigit(rdnString[position]))
                            {
                                WriteUTF8Bytes(rawData, rdnString[position]);
                                position++;
                            }
                            else
                            {
                                throw new ArgumentException("Invalid Attribute OID: OID must start with a digit", rdnString);
                            }

                            #endregion

                            #region The remaining characters need to be a digit or a period

                            while (IsDigit(rdnString[position]) || rdnString[position] == '.')
                            {
                                WriteUTF8Bytes(rawData, rdnString[position]);
                                position++;
                            }

                            #endregion

                            #region The OID can be followed by any number of blank spaces

                            while (rdnString[position] == ' ')
                            {
                                position++;
                            }

                            #endregion

                            #region And it needs to end with an equals sign

                            if (rdnString[position] == '=')
                            {
                                position++;

                                string OID = Encoding.UTF8.GetString(rawData.ToArray());

                                #region OIDs aren't allowed to end with a period

                                if (OID.EndsWith("."))
                                {
                                    throw new ArgumentException("Invalid RDN: OID cannot end with a period", rdnString);
                                }

                                #endregion

                                #region You're also not allowed to have two periods in a row

                                if (OID.IndexOf("..") > -1)
                                    throw new ArgumentException("Invalid RDN: OIDs cannot have two periods in a row", rdnString);

                                #endregion

                                #region Also, numbers aren't allowed to have a leading zero

                                string[] oidPieces = OID.Split('.');

                                foreach (string oidPiece in oidPieces)
                                {
                                    if (oidPiece.Length > 1 && oidPiece[0] == '0')
                                        throw new ArgumentException("Invalid RDN: OIDs cannot have a leading zero", rdnString);
                                }

                                #endregion

                                rawTypes.Add(OID);
                                rawData.SetLength(0);
                                state = ParserState.DetermineValueType;
                            }
                            else
                            {
                                throw new ArgumentException("Invalid RDN: unterminated attribute name", rdnString);
                            }


                            #endregion
                        }
                        catch (IndexOutOfRangeException)
                        {
                            throw new ArgumentException("Invalid RDN: unterminated attribute OID", rdnString);
                        }

                        break;

                    #endregion

                    #region case ParserState.DetermineValueType:

                    case ParserState.DetermineValueType:

                        try
                        {
                            #region Get rid of any leading spaces

                            while (rdnString[position] == ' ')
                            {
                                position++;
                            }

                            #endregion

                            #region Find out what the first character of the value is and set the state accordingly

                            if (rdnString[position] == '"')
                            {
                                state = ParserState.GetQuotedValue;
                            }
                            else if (rdnString[position] == '#')
                            {
                                state = ParserState.GetHexValue;
                            }
                            else
                            {
                                state = ParserState.GetUnquotedValue;
                            }

                            #endregion
                        }
                        catch (IndexOutOfRangeException)
                        {
                            state = ParserState.GetUnquotedValue;
                            rawValues.Add("");
                            rawValueTypes.Add(RDNComponent.RDNValueType.StringValue);
                        }

                        break;

                    #endregion

                    #region case ParserState.GetHexValue:

                    case ParserState.GetHexValue:

                        #region The first character has to be a pound sign

                        if (rdnString[position] == '#')
                        {
                            WriteUTF8Bytes(rawData, rdnString[position]);
                            position++;
                        }
                        else
                        {
                            throw new ArgumentException("Invalid RDN: Hex values must start with #", rdnString);
                        }

                        #endregion

                        #region The rest of the characters have to be hex digits

                        while (position + 1 < rdnString.Length && IsHexDigit(rdnString[position]) && IsHexDigit(rdnString[position + 1]))
                        {
                            WriteUTF8Bytes(rawData, rdnString[position]);
                            WriteUTF8Bytes(rawData, rdnString[position + 1]);
                            position += 2;
                        }

                        #endregion

                        #region Get rid of any trailing blank spaces

                        while (position < rdnString.Length && rdnString[position] == ' ')
                        {
                            position++;
                        }

                        #endregion

                        #region Check for end-of-string or + sign (which indicates a multi-valued RDN)

                        if (position >= rdnString.Length)
                        {
                            rawValues.Add(Encoding.UTF8.GetString(rawData.ToArray()));
                            rawData.SetLength(0);
                            rawValueTypes.Add(RDNComponent.RDNValueType.HexValue);
                        }
                        else if (rdnString[position] == '+')
                        {
                            position++;
                            rawValues.Add(Encoding.UTF8.GetString(rawData.ToArray()));
                            rawData.SetLength(0);
                            state = ParserState.DetermineAttributeType;
                            rawValueTypes.Add(RDNComponent.RDNValueType.HexValue);
                        }
                        else
                        {
                            throw new ArgumentException("Invalid RDN: Invalid characters at end of value", rdnString);
                        }

                        #endregion

                        break;

                    #endregion

                    #region case ParserState.GetQuotedValue:

                    case ParserState.GetQuotedValue:

                        #region The first character has to be a quote

                        if (rdnString[position] == '"')
                        {
                            position++;
                        }
                        else
                        {
                            throw new ArgumentException("Invalid RDN: Quoted values must start with \"", rdnString);
                        }

                        #endregion

                        #region Read characters until we find a closing quote

                        try
                        {
                            while (rdnString[position] != '"')
                            {
                                if (rdnString[position] == '\\')
                                {
                                    try
                                    {
                                        position++;

                                        if (IsHexDigit(rdnString[position]))
                                        {
                                            rawData.WriteByte(Convert.ToByte(rdnString.Substring(position, 2), 16));

                                            position += 2;
                                        }
                                        else if (IsSpecialChar(rdnString[position]) || rdnString[position] == ' ')
                                        {
                                            WriteUTF8Bytes(rawData, rdnString[position]);
                                            position++;
                                        }
                                        else
                                        {
                                            throw new ArgumentException("Invalid RDN: Escape sequence \\" + rdnString[position] + " is invalid.", rdnString);
                                        }
                                    }
                                    catch (IndexOutOfRangeException)
                                    {
                                        throw new ArgumentException("Invalid RDN: Invalid escape sequence", rdnString);
                                    }
                                }
                                else
                                {
                                    WriteUTF8Bytes(rawData, rdnString[position]);
                                    position++;
                                }
                            }
                        }
                        catch (IndexOutOfRangeException)
                        {
                            throw new ArgumentException("Invalid RDN: Unterminated quoted value", rdnString);
                        }

                        position++;

                        #endregion

                        #region Remove any trailing spaces

                        while (position < rdnString.Length && rdnString[position] == ' ')
                        {
                            position++;
                        }

                        #endregion

                        #region Check for end-of-string or + sign (which indicates a multi-valued RDN)

                        if (position >= rdnString.Length)
                        {
                            rawValues.Add(Encoding.UTF8.GetString(rawData.ToArray()));
                            rawData.SetLength(0);
                            rawValueTypes.Add(RDNComponent.RDNValueType.StringValue);
                        }
                        else if (rdnString[position] == '+')
                        {
                            rawValues.Add(Encoding.UTF8.GetString(rawData.ToArray()));
                            rawData.SetLength(0);
                            state = ParserState.DetermineAttributeType;
                            rawValueTypes.Add(RDNComponent.RDNValueType.StringValue);
                        }
                        else
                        {
                            throw new ArgumentException("Invalid RDN: Invalid characters at end of value", rdnString);
                        }

                        #endregion

                        break;

                    #endregion

                    #region case ParserState.GetUnquotedValue:

                    case ParserState.GetUnquotedValue:

                        #region Is the first character an escaped space or pound sign?

                        try
                        {
                            if (rdnString[position] == '\\' && (rdnString[position + 1] == ' ' || rdnString[position + 1] == '#'))
                            {
                                WriteUTF8Bytes(rawData, rdnString[position + 1]);
                                position += 2;
                            }
                        }
                        catch (IndexOutOfRangeException)
                        {
                            throw new ArgumentException("Invalid RDN: Invalid escape sequence", rdnString);
                        }

                        #endregion

                        #region Read all characters, keeping track of trailing spaces

                        int trailingSpaces = 0;

                        while (position < rdnString.Length && rdnString[position] != '+')
                        {
                            if (rdnString[position] == ' ')
                            {
                                trailingSpaces++;

                                WriteUTF8Bytes(rawData, rdnString[position]);

                                position++;
                            }
                            else
                            {
                                trailingSpaces = 0;

                                if (rdnString[position] == '\\')
                                {
                                    try
                                    {
                                        position++;

                                        if (IsHexDigit(rdnString[position]))
                                        {
                                            rawData.WriteByte(Convert.ToByte(rdnString.Substring(position, 2), 16));

                                            position += 2;
                                        }
                                        else if (IsSpecialChar(rdnString[position]) || rdnString[position] == ' ')
                                        {
                                            WriteUTF8Bytes(rawData, rdnString[position]);
                                            position++;
                                        }
                                        else
                                        {
                                            throw new ArgumentException("Invalid RDN: Escape sequence \\" + rdnString[position] + " is invalid.", rdnString);
                                        }
                                    }
                                    catch (IndexOutOfRangeException)
                                    {
                                        throw new ArgumentException("Invalid RDN: Invalid escape sequence", rdnString);
                                    }
                                }
                                else if (IsSpecialChar(rdnString[position]))
                                {
                                    throw new ArgumentException("Invalid RDN: Unquoted special character '" + rdnString[position] + "'", rdnString);
                                }
                                else
                                {
                                    WriteUTF8Bytes(rawData, rdnString[position]);
                                    position++;
                                }
                            }
                        }

                        #endregion

                        #region Remove any trailing spaces

                        rawData.SetLength(rawData.Length - trailingSpaces);

                        #endregion

                        #region If the last character is a + sign, set state to look for another value, otherwise, finish up

                        if (position < rdnString.Length)
                        {
                            if (rdnString[position] == '+')
                            {
                                position++;
                                rawValues.Add(Encoding.UTF8.GetString(rawData.ToArray()));
                                rawData.SetLength(0);
                                state = ParserState.DetermineAttributeType;
                                rawValueTypes.Add(RDNComponent.RDNValueType.StringValue);
                            }
                            else
                            {
                                throw new ArgumentException("Invalid RDN: invalid trailing character", rdnString);
                            }
                        }
                        else
                        {
                            rawValues.Add(Encoding.UTF8.GetString(rawData.ToArray()));
                            rawData.SetLength(0);
                            rawValueTypes.Add(RDNComponent.RDNValueType.StringValue);
                        }

                        #endregion

                        break;

                        #endregion
                }
            }

            #region Check ending state

            if (!(state == ParserState.GetHexValue ||
                state == ParserState.GetQuotedValue ||
                state == ParserState.GetUnquotedValue
                ))
            {
                throw new ArgumentException("Invalid RDN", rdnString);
            }

            #endregion

            #region Store the results we've collected

            RDNComponent[] componentArray = new RDNComponent[rawTypes.Count];

            for (int i = 0; i < componentArray.Length; i++)
            {
                componentArray[i] = new RDNComponent((string)rawTypes[i], (string)rawValues[i], (RDNComponent.RDNValueType)rawValueTypes[i]);
            }

            components = new RDNComponentList(componentArray);

            #endregion
        }


        #endregion

        #region Static Methods

        public static bool IsAlpha(char c)
        {
            return (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z');
        }

        public static bool IsDigit(char c)
        {
            return (c >= '0' && c <= '9');
        }

        public static bool IsHexDigit(char c)
        {
            return (c >= '0' && c <= '9') || (c >= 'A' && c <= 'F') || (c >= 'a' && c <= 'f');
        }

        public static bool IsSpecialChar(char c)
        {
            return c == ',' || c == '=' || c == '+' || c == '<' || c == '>' || c == '#' || c == ';' || c == '\\' || c == '"';
        }

        #endregion

        #region Overloaded Operators

        public static bool operator ==(RDN obj1, RDN obj2)
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

        public static bool operator !=(RDN obj1, RDN obj2)
        {
            return (!(obj1 == obj2));
        }

        #endregion
    }

    public class RDNList : IEnumerable
    {
        #region Data Members

        private RDN[] rDNs;

        #endregion

        #region Constructors

        internal RDNList(RDN[] rDNs)
        {
            this.rDNs = rDNs;
        }

        #endregion

        #region Properties

        public int Length
        {
            get
            {
                return rDNs.Length;
            }
        }

        public RDN this[int index]
        {
            get
            {
                return rDNs[index];
            }
        }

        #endregion

        #region IEnumerable Members

        public IEnumerator GetEnumerator()
        {
            return rDNs.GetEnumerator();
        }

        #endregion
    }
}
