// ConvWab
// Many items are brought from libwab and wabread which had been published at filewut.com

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Xml.Linq;

namespace ConvWab {
    class Program {
        static void Main(string[] args) {
            if (args.Length < 2) {
                Console.Error.WriteLine("ConvWab <in.wab> <out.xml>");
                Environment.ExitCode = 1;
            }
            else {
                new Program().Run(args[0], args[1]);
            }
        }

        private void Run(string fp, string fpxml) {
            MemoryStream si = new MemoryStream(File.ReadAllBytes(fp));
            wab_handle wh = new wab_handle();
            wh.fp = si;
            BinaryReader br = new BinaryReader(si);
            wh.wabhead = new wab_header(br);
            XElement elwab;
            XDocument xmlo = new XDocument(
                elwab = new XElement("wab")
                );
            int z = -1;
            foreach (tabledesc td in wh.wabhead.tables) {
                XElement eltbl;
                z++;
                elwab.Add(
                    eltbl = new XElement("table"
                        , new XAttribute("i", z)
                        )
                    );

                if (td.IsIdx) {
                    eltbl.SetAttributeValue("t", "index");
                    int y = -1;
                    foreach (idxrecord idx in td.midx) {
                        y++;
                        XElement elrec;
                        eltbl.Add(
                            elrec = new XElement("record"
                                , new XAttribute("i", y)
                                )
                            );
                        si.Position = idx.offset;
                        wab_record rec = new wab_record(br);
                        foreach (var sr in rec.srecs) {
                            XElement els;
                            elrec.Add(
                                els = new XElement("field"
                                    , new XAttribute("t", sr.ElemType)
                                    , new XAttribute("idHex", String.Format("{0:X4}", (sr.opcode >> 16) & 0xFFFF))
                                    , new XAttribute("opcHex", String.Format("{0:X8}", (sr.opcode)))
                                    , new XAttribute("isArray", sr.IsArray)
                                    , new XAttribute("id", String.Format("{0}", (sr.opcode >> 16) & 0xFFFF))
                                    , new XAttribute("opc", String.Format("{0}", (sr.opcode)))
                                    , new XAttribute("ldid", UtPR.Ldid((sr.opcode >> 16) & 0xFFFF) ?? "")
                                    , new XAttribute("prop", UtPR.Id((sr.opcode >> 16) & 0xFFFF) ?? "")
                                    )
                                );

                            Object[] al = new Object[1] { sr.Render };
                            if (sr.IsArray) {
                                al = new System.Collections.ArrayList(((Array)al[0])).ToArray();
                            }
                            int x = -1;
                            foreach (Object item in al) {
                                x++;
                                els.Add(
                                    new XElement("item"
                                        , new XAttribute("i", x)
                                        , StrSerialize(item)
                                        )
                                    );
                            }
                        }
                    }
                }
                else if (td.IsTxt) {
                    eltbl.SetAttributeValue("t", "text");

                    foreach (txtrecord rec in td.mtxt) {
                        eltbl.Add(
                            new XElement("text"
                                , new XAttribute("id", rec.recid)
                                , new XText(Encoding.Unicode.GetString(rec.str).Split('\0')[0])
                                )
                            );
                    }
                }
            }

            xmlo.Save(fpxml);
        }

        class UtPR {
            public static string Ldid(int id) {
                switch (id) {
                    case PR_DISPLAY_NAME: return "dc";
                    case PR_MAB_ADDRESS_STR: return "mail";
                    case PR_COMMENT_STR: return "ntUserComment";
                    case PR_GIVEN_NAME_STR: return "givenName";
                    case PR_MAB_BUSINESS_PHONE: return "telephoneNumber";
                    case PR_MAB_PHONE: return "homePhone";
                    case PR_INITIALS_STR: return "initials";
                    case PR_SURNAME_STR: return "sn";
                    case PR_MAB_COMPANY: return "o";
                    case PR_MAB_JOBTITLE: return "title";
                    case PR_MAB_DEPARTMENT: return "ou";
                    case PR_MAB_OFFICE: return "ou";
                    case PR_MAB_MOBILE: return "mobile";
                    case PR_MAB_PAGER: return "pager";
                    case PR_MAB_BUSINESS_FAX: return "fax";
                    case PR_MAB_FAX: return "fax";
                    case PR_MAB_BUSINESS_COUNTRY: return "businessCountry";
                    case PR_MAB_BUSINESS_CITY: return "businessLocality";
                    case PR_MAB_BUSINESS_PROVINCE: return "businessSt";
                    case PR_MAB_BUSINESS_ADDRESS: return "businessStreet";
                    case PR_MAB_BUSINESS_POSTAL_CODE: return "businessPostal";
                    case PR_MAB_BUSINESS_MIDDLE: return "businessMiddle";
                    case PR_MAB_BUSINESS_TITLE: return "title";
                    case PR_MAB_NICK: return "nickName";
                    case PR_MAB_BUSINESS_URL1: return "businessUrl";
                    case PR_MAB_BUSINESS_URL2: return "businessUrl";
                    case PR_MAB_ALTERNATE_EMAILS: return "mail";
                    case PR_MAB_CITY: return "l";
                    case PR_MAB_COUNTRY: return "c";
                    case PR_MAB_POSTAL_CODE: return "postalCode";
                    case PR_MAB_PROVINCE: return "st";
                    case PR_MAB_STREET_ADDRESS: return "street";
                    case PR_MAB_MEMBER: return "member";
                    case PR_MAB_IP_PHONE: return "ipPhone";
                    case PR_MAB_EMAIL_ADDRESS_WTF: return "mail";
                }
                return null;
            }

            public static string Id(int id) {
                switch (id) {
                    case PR_ALTERNATE_RECIPIENT_ALLOWED_STR: return "PR_ALTERNATE_RECIPIENT_ALLOWED_STR";
                    case PR_IMPORTANCE_STR: return "PR_IMPORTANCE_STR";
                    case PR_MESSAGE_CLASS_STR: return "PR_MESSAGE_CLASS_STR";
                    case PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED_STR: return "PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED_STR";
                    case PR_PRIORITY_STR: return "PR_PRIORITY_STR";
                    case PR_READ_RECEIPT_REQUESTED_STR: return "PR_READ_RECEIPT_REQUESTED_STR";
                    case PR_ORIGINAL_SENSITIVITY_STR: return "PR_ORIGINAL_SENSITIVITY_STR";
                    case PR_SENSITIVITY_STR: return "PR_SENSITIVITY_STR";
                    case PR_SUBJECT_STR: return "PR_SUBJECT_STR";
                    case PR_CLIENT_SUBMIT_TIME_STR: return "PR_CLIENT_SUBMIT_TIME_STR";
                    case PR_SENT_REPRESENTING_SEARCH_KEY_STR: return "PR_SENT_REPRESENTING_SEARCH_KEY_STR";
                    case PR_RECEIVED_BY_ENTRYID_STR: return "PR_RECEIVED_BY_ENTRYID_STR";
                    case PR_RECEIVED_BY_NAME_STR: return "PR_RECEIVED_BY_NAME_STR";
                    case PR_SENT_REPRESENTING_ENTRYID_STR: return "PR_SENT_REPRESENTING_ENTRYID_STR";
                    case PR_SENT_REPRESENTING_NAME_STR: return "PR_SENT_REPRESENTING_NAME_STR";
                    case PR_RCVD_REPRESENTING_ENTRYID_STR: return "PR_RCVD_REPRESENTING_ENTRYID_STR";
                    case PR_RCVD_REPRESENTING_NAME_STR: return "PR_RCVD_REPRESENTING_NAME_STR";
                    case PR_REPLY_RECIPIENT_ENTRIES_STR: return "PR_REPLY_RECIPIENT_ENTRIES_STR";
                    case PR_REPLY_RECIPIENT_NAMES_STR: return "PR_REPLY_RECIPIENT_NAMES_STR";
                    case PR_RECEIVED_BY_SEARCH_KEY_STR: return "PR_RECEIVED_BY_SEARCH_KEY_STR";
                    case PR_RCVD_REPRESENTING_SEARCH_KEY_STR: return "PR_RCVD_REPRESENTING_SEARCH_KEY_STR";
                    case PR_MESSAGE_TO_ME_STR: return "PR_MESSAGE_TO_ME_STR";
                    case PR_MESSAGE_CC_ME_STR: return "PR_MESSAGE_CC_ME_STR";
                    case PR_MESSAGE_RECIP_ME_STR: return "PR_MESSAGE_RECIP_ME_STR";
                    case PR_SENT_REPRESENTING_ADDRTYPE_STR: return "PR_SENT_REPRESENTING_ADDRTYPE_STR";
                    case PR_SENT_REPRESENTING_EMAIL_ADDRESS_STR: return "PR_SENT_REPRESENTING_EMAIL_ADDRESS_STR";
                    case PR_CONVERSATION_TOPIC_STR: return "PR_CONVERSATION_TOPIC_STR";
                    case PR_CONVERSATION_INDEX_STR: return "PR_CONVERSATION_INDEX_STR";
                    case PR_RECEIVED_BY_ADDRTYPE_STR: return "PR_RECEIVED_BY_ADDRTYPE_STR";
                    case PR_RECEIVED_BY_EMAIL_ADDRESS_STR: return "PR_RECEIVED_BY_EMAIL_ADDRESS_STR";
                    case PR_RCVD_REPRESENTING_ADDRTYPE_STR: return "PR_RCVD_REPRESENTING_ADDRTYPE_STR";
                    case PR_RCVD_REPRESENTING_EMAIL_ADDRESS_STR: return "PR_RCVD_REPRESENTING_EMAIL_ADDRESS_STR";
                    case PR_TRANSPORT_MESSAGE_HEADERS_STR: return "PR_TRANSPORT_MESSAGE_HEADERS_STR";
                    case PR_SENDER_ENTRYID_STR: return "PR_SENDER_ENTRYID_STR";
                    case PR_SENDER_NAME_STR: return "PR_SENDER_NAME_STR";
                    case PR_SENDER_SEARCH_KEY_STR: return "PR_SENDER_SEARCH_KEY_STR";
                    case PR_SENDER_ADDRTYPE_STR: return "PR_SENDER_ADDRTYPE_STR";
                    case PR_SENDER_EMAIL_ADDRESS_STR: return "PR_SENDER_EMAIL_ADDRESS_STR";
                    case PR_DELETE_AFTER_SUBMIT_STR: return "PR_DELETE_AFTER_SUBMIT_STR";
                    case PR_DISPLAY_CC_STR: return "PR_DISPLAY_CC_STR";
                    case PR_DISPLAY_TO_STR: return "PR_DISPLAY_TO_STR";
                    case PR_MESSAGE_DELIVERY_TIME_STR: return "PR_MESSAGE_DELIVERY_TIME_STR";
                    case PR_MESSAGE_FLAGS_STR: return "PR_MESSAGE_FLAGS_STR";
                    case PR_MESSAGE_SIZE_STR: return "PR_MESSAGE_SIZE_STR";
                    case PR_SENTMAIL_ENTRYID_STR: return "PR_SENTMAIL_ENTRYID_STR";
                    case PR_RTF_IN_SYNC_STR: return "PR_RTF_IN_SYNC_STR";
                    case PR_ATTACH_SIZE_STR: return "PR_ATTACH_SIZE_STR";
                    case PR_RECORD_KEY_STR: return "PR_RECORD_KEY_STR";
                    case PR_MAB_MYSTERY_ALWAYS_6: return "PR_MAB_MYSTERY_ALWAYS_6";
                    case PR_MAB_ENTRY_ID: return "PR_MAB_ENTRY_ID";
                    case PR_BODY_STR: return "PR_BODY_STR";
                    case PR_RTF_SYNC_BODY_CRC_STR: return "PR_RTF_SYNC_BODY_CRC_STR";
                    case PR_RTF_SYNC_BODY_COUNT_STR: return "PR_RTF_SYNC_BODY_COUNT_STR";
                    case PR_RTF_SYNC_BODY_TAG_STR: return "PR_RTF_SYNC_BODY_TAG_STR";
                    case PR_RTF_COMPRESSED_STR: return "PR_RTF_COMPRESSED_STR";
                    case PR_RTF_SYNC_PREFIX_COUNT_STR: return "PR_RTF_SYNC_PREFIX_COUNT_STR";
                    case PR_RTF_SYNC_TRAILING_COUNT_STR: return "PR_RTF_SYNC_TRAILING_COUNT_STR";
                    case PR_HTML_BODY_STR: return "PR_HTML_BODY_STR";
                    case PR_MESSAGE_ID_STR: return "PR_MESSAGE_ID_STR";
                    case PR_IN_REPLY_TO_STR: return "PR_IN_REPLY_TO_STR";
                    case PR_RETURN_PATH_STR: return "PR_RETURN_PATH_STR";
                    case PR_DISPLAY_NAME: return "PR_DISPLAY_NAME";
                    case PR_MAB_SEND_METHOD: return "PR_MAB_SEND_METHOD";
                    case PR_MAB_ADDRESS_STR: return "PR_MAB_ADDRESS_STR";
                    case PR_COMMENT_STR: return "PR_COMMENT_STR";
                    case PR_CREATION_TIME_STR: return "PR_CREATION_TIME_STR";
                    case PR_LAST_MODIFICATION_TIME_STR: return "PR_LAST_MODIFICATION_TIME_STR";
                    case PR_SEARCH_KEY_STR: return "PR_SEARCH_KEY_STR";
                    case PR_VALID_FOLDER_MASK_STR: return "PR_VALID_FOLDER_MASK_STR";
                    case PR_IPM_SUBTREE_ENTRYID_STR: return "PR_IPM_SUBTREE_ENTRYID_STR";
                    case PR_IPM_WASTEBASKET_ENTRYID_STR: return "PR_IPM_WASTEBASKET_ENTRYID_STR";
                    case PR_FINDER_ENTRYID_STR: return "PR_FINDER_ENTRYID_STR";
                    case PR_CONTENT_COUNT_STR: return "PR_CONTENT_COUNT_STR";
                    case PR_CONTENT_UNREAD_STR: return "PR_CONTENT_UNREAD_STR";
                    case PR_SUBFOLDERS_STR: return "PR_SUBFOLDERS_STR";
                    case PR_CONTAINER_CLASS_STR: return "PR_CONTAINER_CLASS_STR";
                    case PR_ASSOC_CONTENT_COUNT_STR: return "PR_ASSOC_CONTENT_COUNT_STR";
                    case PR_ATTACH_DATA_OBJ_STR: return "PR_ATTACH_DATA_OBJ_STR";
                    case PR_ATTACH_FILENAME_STR: return "PR_ATTACH_FILENAME_STR";
                    case PR_ATTACH_METHOD_STR: return "PR_ATTACH_METHOD_STR";
                    case PR_ATTACH_LONG_FILENAME_STR: return "PR_ATTACH_LONG_FILENAME_STR";
                    case PR_RENDERING_POSITION_STR: return "PR_RENDERING_POSITION_STR";
                    case PR_ATTACH_MIME_TAG_STR: return "PR_ATTACH_MIME_TAG_STR";
                    case PR_ATTACH_MIME_SEQUENCE_STR: return "PR_ATTACH_MIME_SEQUENCE_STR";
                    case PR_GIVEN_NAME_STR: return "PR_GIVEN_NAME_STR";
                    case PR_MAB_BUSINESS_PHONE: return "PR_MAB_BUSINESS_PHONE";
                    case PR_MAB_PHONE: return "PR_MAB_PHONE";
                    case PR_INITIALS_STR: return "PR_INITIALS_STR";
                    case PR_SURNAME_STR: return "PR_SURNAME_STR";
                    case PR_MAB_COMPANY: return "PR_MAB_COMPANY";
                    case PR_MAB_JOBTITLE: return "PR_MAB_JOBTITLE";
                    case PR_MAB_DEPARTMENT: return "PR_MAB_DEPARTMENT";
                    case PR_MAB_OFFICE: return "PR_MAB_OFFICE";
                    case PR_MAB_MOBILE: return "PR_MAB_MOBILE";
                    case PR_MAB_PAGER: return "PR_MAB_PAGER";
                    case PR_MAB_BUSINESS_FAX: return "PR_MAB_BUSINESS_FAX";
                    case PR_MAB_FAX: return "PR_MAB_FAX";
                    case PR_MAB_BUSINESS_COUNTRY: return "PR_MAB_BUSINESS_COUNTRY";
                    case PR_MAB_BUSINESS_CITY: return "PR_MAB_BUSINESS_CITY";
                    case PR_MAB_BUSINESS_PROVINCE: return "PR_MAB_BUSINESS_PROVINCE";
                    case PR_MAB_BUSINESS_ADDRESS: return "PR_MAB_BUSINESS_ADDRESS";
                    case PR_MAB_BUSINESS_POSTAL_CODE: return "PR_MAB_BUSINESS_POSTAL_CODE";
                    case PR_MAB_BUSINESS_MIDDLE: return "PR_MAB_BUSINESS_MIDDLE";
                    case PR_MAB_BUSINESS_TITLE: return "PR_MAB_BUSINESS_TITLE";
                    case PR_MAB_NICK: return "PR_MAB_NICK";
                    case PR_MAB_BUSINESS_URL1: return "PR_MAB_BUSINESS_URL1";
                    case PR_MAB_BUSINESS_URL2: return "PR_MAB_BUSINESS_URL2";
                    case PR_MAB_ALTERNATE_EMAIL_TYPES: return "PR_MAB_ALTERNATE_EMAIL_TYPES";
                    case PR_MAB_MYSTERY_ALWAYS_0_FIRST: return "PR_MAB_MYSTERY_ALWAYS_0_FIRST";
                    case PR_MAB_ALTERNATE_EMAILS: return "PR_MAB_ALTERNATE_EMAILS";
                    case PR_MAB_CITY: return "PR_MAB_CITY";
                    case PR_MAB_COUNTRY: return "PR_MAB_COUNTRY";
                    case PR_MAB_POSTAL_CODE: return "PR_MAB_POSTAL_CODE";
                    case PR_MAB_PROVINCE: return "PR_MAB_PROVINCE";
                    case PR_MAB_STREET_ADDRESS: return "PR_MAB_STREET_ADDRESS";
                    case PR_MAB_MYSTERY_ALWAYS_1048576: return "PR_MAB_MYSTERY_ALWAYS_1048576";
                    case PR_MAB_MYSTERY_PROFILE_ID1: return "PR_MAB_MYSTERY_PROFILE_ID1";
                    case PR_MAB_MEMBER: return "PR_MAB_MEMBER";
                    case PR_MAB_IP_PHONE: return "PR_MAB_IP_PHONE";
                    case PR_MAB_MYSTERY_PROFILE_ID2: return "PR_MAB_MYSTERY_PROFILE_ID2";
                    case PR_MAB_EMAIL_ADDRESS_WTF: return "PR_MAB_EMAIL_ADDRESS_WTF";
                    case PR_Schedule_Folder_EntryID_STR: return "PR_Schedule_Folder_EntryID_STR";
                    case PR_ID2_STR: return "PR_ID2_STR";
                    case PR_EXTRA_PROPERTY_IDENTIFIER_STR: return "PR_EXTRA_PROPERTY_IDENTIFIER_STR";
                    case PR_MAB_MYSTERY_ALWAYS_0_NEXT: return "PR_MAB_MYSTERY_ALWAYS_0_NEXT";
                    case PR_ADDRESS_1_STR: return "PR_ADDRESS_1_STR";
                    case PR_ACCESS_METHOD_STR: return "PR_ACCESS_METHOD_STR";
                    case PR_ADDRESS_1_DESCRIPTION_STR: return "PR_ADDRESS_1_DESCRIPTION_STR";
                }
                return null;
            }

            const int PR_ALTERNATE_RECIPIENT_ALLOWED_STR = 0x0002;
            const int PR_IMPORTANCE_STR = 0x0017;
            const int PR_MESSAGE_CLASS_STR = 0x001A;
            const int PR_ORIGINATOR_DELIVERY_REPORT_REQUESTED_STR = 0x0023;
            const int PR_PRIORITY_STR = 0x0026;
            const int PR_READ_RECEIPT_REQUESTED_STR = 0x0029;
            const int PR_ORIGINAL_SENSITIVITY_STR = 0x002E;
            const int PR_SENSITIVITY_STR = 0x0036;
            const int PR_SUBJECT_STR = 0x0037;
            const int PR_CLIENT_SUBMIT_TIME_STR = 0x0039;
            const int PR_SENT_REPRESENTING_SEARCH_KEY_STR = 0x003B;
            const int PR_RECEIVED_BY_ENTRYID_STR = 0x003F;
            const int PR_RECEIVED_BY_NAME_STR = 0x0040;
            const int PR_SENT_REPRESENTING_ENTRYID_STR = 0x0041;
            const int PR_SENT_REPRESENTING_NAME_STR = 0x0042;
            const int PR_RCVD_REPRESENTING_ENTRYID_STR = 0x0043;
            const int PR_RCVD_REPRESENTING_NAME_STR = 0x0044;
            const int PR_REPLY_RECIPIENT_ENTRIES_STR = 0x004F;
            const int PR_REPLY_RECIPIENT_NAMES_STR = 0x0050;
            const int PR_RECEIVED_BY_SEARCH_KEY_STR = 0x0051;
            const int PR_RCVD_REPRESENTING_SEARCH_KEY_STR = 0x0052;
            const int PR_MESSAGE_TO_ME_STR = 0x0057;
            const int PR_MESSAGE_CC_ME_STR = 0x0058;
            const int PR_MESSAGE_RECIP_ME_STR = 0x0059;
            const int PR_SENT_REPRESENTING_ADDRTYPE_STR = 0x0064;
            const int PR_SENT_REPRESENTING_EMAIL_ADDRESS_STR = 0x0065;
            const int PR_CONVERSATION_TOPIC_STR = 0x0070;
            const int PR_CONVERSATION_INDEX_STR = 0x0071;
            const int PR_RECEIVED_BY_ADDRTYPE_STR = 0x0075;
            const int PR_RECEIVED_BY_EMAIL_ADDRESS_STR = 0x0076;
            const int PR_RCVD_REPRESENTING_ADDRTYPE_STR = 0x0077;
            const int PR_RCVD_REPRESENTING_EMAIL_ADDRESS_STR = 0x0078;
            const int PR_TRANSPORT_MESSAGE_HEADERS_STR = 0x007D;
            const int PR_SENDER_ENTRYID_STR = 0x0C19;
            const int PR_SENDER_NAME_STR = 0x0C1A;
            const int PR_SENDER_SEARCH_KEY_STR = 0x0C1D;
            const int PR_SENDER_ADDRTYPE_STR = 0x0C1E;
            const int PR_SENDER_EMAIL_ADDRESS_STR = 0x0C1F;
            const int PR_DELETE_AFTER_SUBMIT_STR = 0x0E01;
            const int PR_DISPLAY_CC_STR = 0x0E03;
            const int PR_DISPLAY_TO_STR = 0x0E04;
            const int PR_MESSAGE_DELIVERY_TIME_STR = 0x0E06;
            const int PR_MESSAGE_FLAGS_STR = 0x0E07;
            const int PR_MESSAGE_SIZE_STR = 0x0E08;
            const int PR_SENTMAIL_ENTRYID_STR = 0x0E0A;
            const int PR_RTF_IN_SYNC_STR = 0x0E1F;
            const int PR_ATTACH_SIZE_STR = 0x0E20;
            const int PR_RECORD_KEY_STR = 0x0FF9;
            const int PR_MAB_MYSTERY_ALWAYS_6 = 0x0ffe;
            const int PR_MAB_ENTRY_ID = 0x0fff;
            const int PR_BODY_STR = 0x1000;
            const int PR_RTF_SYNC_BODY_CRC_STR = 0x1006;
            const int PR_RTF_SYNC_BODY_COUNT_STR = 0x1007;
            const int PR_RTF_SYNC_BODY_TAG_STR = 0x1008;
            const int PR_RTF_COMPRESSED_STR = 0x1009;
            const int PR_RTF_SYNC_PREFIX_COUNT_STR = 0x1010;
            const int PR_RTF_SYNC_TRAILING_COUNT_STR = 0x1011;
            const int PR_HTML_BODY_STR = 0x1013;
            const int PR_MESSAGE_ID_STR = 0x1035;
            const int PR_IN_REPLY_TO_STR = 0x1042;
            const int PR_RETURN_PATH_STR = 0x1046;
            const int PR_DISPLAY_NAME = 0x3001;
            const int PR_MAB_SEND_METHOD = 0x3002;
            const int PR_MAB_ADDRESS_STR = 0x3003;
            const int PR_COMMENT_STR = 0x3004;
            const int PR_CREATION_TIME_STR = 0x3007;
            const int PR_LAST_MODIFICATION_TIME_STR = 0x3008;
            const int PR_SEARCH_KEY_STR = 0x300B;
            const int PR_VALID_FOLDER_MASK_STR = 0x35DF;
            const int PR_IPM_SUBTREE_ENTRYID_STR = 0x35E0;
            const int PR_IPM_WASTEBASKET_ENTRYID_STR = 0x35E3;
            const int PR_FINDER_ENTRYID_STR = 0x35E7;
            const int PR_CONTENT_COUNT_STR = 0x3602;
            const int PR_CONTENT_UNREAD_STR = 0x3603;
            const int PR_SUBFOLDERS_STR = 0x360A;
            const int PR_CONTAINER_CLASS_STR = 0x3613;
            const int PR_ASSOC_CONTENT_COUNT_STR = 0x3617;
            const int PR_ATTACH_DATA_OBJ_STR = 0x3701;
            const int PR_ATTACH_FILENAME_STR = 0x3704;
            const int PR_ATTACH_METHOD_STR = 0x3705;
            const int PR_ATTACH_LONG_FILENAME_STR = 0x3707;
            const int PR_RENDERING_POSITION_STR = 0x370B;
            const int PR_ATTACH_MIME_TAG_STR = 0x370E;
            const int PR_ATTACH_MIME_SEQUENCE_STR = 0x3710;
            const int PR_GIVEN_NAME_STR = 0x3a06;
            const int PR_MAB_BUSINESS_PHONE = 0x3a08;
            const int PR_MAB_PHONE = 0x3a09;
            const int PR_INITIALS_STR = 0x3a0a;
            const int PR_SURNAME_STR = 0x3a11;
            const int PR_MAB_COMPANY = 0x3a16;
            const int PR_MAB_JOBTITLE = 0x3a17;
            const int PR_MAB_DEPARTMENT = 0x3a18;
            const int PR_MAB_OFFICE = 0x3a19;
            const int PR_MAB_MOBILE = 0x3a1c;
            const int PR_MAB_PAGER = 0x3a21;
            const int PR_MAB_BUSINESS_FAX = 0x3a24;
            const int PR_MAB_FAX = 0x3a25;
            const int PR_MAB_BUSINESS_COUNTRY = 0x3a26;
            const int PR_MAB_BUSINESS_CITY = 0x3a27;
            const int PR_MAB_BUSINESS_PROVINCE = 0x3a28;
            const int PR_MAB_BUSINESS_ADDRESS = 0x3a29;
            const int PR_MAB_BUSINESS_POSTAL_CODE = 0x3a2a;
            const int PR_MAB_BUSINESS_MIDDLE = 0x3a44;
            const int PR_MAB_BUSINESS_TITLE = 0x3a45;
            const int PR_MAB_NICK = 0x3a4f;
            const int PR_MAB_BUSINESS_URL1 = 0x3a50;
            const int PR_MAB_BUSINESS_URL2 = 0x3a51;
            const int PR_MAB_ALTERNATE_EMAIL_TYPES = 0x3a54;
            const int PR_MAB_MYSTERY_ALWAYS_0_FIRST = 0x3a55;
            const int PR_MAB_ALTERNATE_EMAILS = 0x3a56;
            const int PR_MAB_CITY = 0x3a59;
            const int PR_MAB_COUNTRY = 0x3a5a;
            const int PR_MAB_POSTAL_CODE = 0x3a5b;
            const int PR_MAB_PROVINCE = 0x3a5c;
            const int PR_MAB_STREET_ADDRESS = 0x3a5d;
            const int PR_MAB_MYSTERY_ALWAYS_1048576 = 0x3a71;
            const int PR_MAB_MYSTERY_PROFILE_ID1 = 0x8004;
            const int PR_MAB_MEMBER = 0x8009;
            const int PR_MAB_IP_PHONE = 0x800a;
            const int PR_MAB_MYSTERY_PROFILE_ID2 = 0x800d;
            const int PR_MAB_EMAIL_ADDRESS_WTF = 0x8012;
            const int PR_Schedule_Folder_EntryID_STR = 0x661E;
            const int PR_ID2_STR = 0x67F2;
            const int PR_EXTRA_PROPERTY_IDENTIFIER_STR = 0x67FF;
            const int PR_MAB_MYSTERY_ALWAYS_0_NEXT = 0x800c;
            const int PR_ADDRESS_1_STR = 0x8019;
            const int PR_ACCESS_METHOD_STR = 0x80cf;
            const int PR_ADDRESS_1_DESCRIPTION_STR = 0x80d1;
        }

        private XNode StrSerialize(object item) {
            if (item is byte[]) {
                String s = "";
                foreach (byte b in ((byte[])item)) {
                    if (s.Length != 0)
                        s += " ";
                    s += b.ToString("x2");
                }
                return new XText(s);
            }
            return new XText("" + item);
        }
    }

    public class tabledesc {
        public int type, size, offset, count;

        public override string ToString() {
            return string.Format("type={0} size={1} offset={2} count={3}"
                , type, size, offset, count);
        }

        public tabledesc(BinaryReader br) {
            type = br.ReadInt32();
            size = br.ReadInt32();
            offset = br.ReadInt32();
            count = br.ReadInt32();

            Int64 prev = br.BaseStream.Position;
            if (IsIdx) {
                midx = new idxrecord[count];
                br.BaseStream.Position = offset;
                for (int x = 0; x < count; x++) {
                    midx[x] = new idxrecord(br);
                }
            }
            else if (IsTxt) {
                mtxt = new txtrecord[count];
                br.BaseStream.Position = offset;
                for (int x = 0; x < count; x++) {
                    mtxt[x] = new txtrecord(br);
                }
            }
            br.BaseStream.Position = prev;
        }

        public bool IsIdx { get { return (type == TYPE_IDX); } }
        public bool IsTxt { get { return (type == TYPE_TXT); } }

        public idxrecord[] midx = null;
        public txtrecord[] mtxt = null;

        public const int TYPE_IDX = 4000;   //FA0
        public const int TYPE_TXT = 34000; //84D0
    }
    public class wab_header {
        public string magic;
        public int count1;
        public int count2;

        public override string ToString() {
            return string.Format("{0}, {1}", count1, count2);
        }

        public const int TABLE_COUNT = 6;
        public tabledesc[] tables = new tabledesc[TABLE_COUNT];

        public wab_header(BinaryReader br) {
            magic = Encoding.ASCII.GetString(br.ReadBytes(16));
            count1 = br.ReadInt32();
            count2 = br.ReadInt32();
            for (int x = 0; x < TABLE_COUNT; x++) tables[x] = new tabledesc(br);
        }
    }
    public class wab_handle {
        public MemoryStream fp;
        public wab_header wabhead;
    }

    public class idxrecord {
        public uint recid;
        public int offset;

        public idxrecord(BinaryReader br) {
            recid = br.ReadUInt32();
            offset = br.ReadInt32();
        }

        public override string ToString() {
            return string.Format("idx {0} {1}", recid, offset);
        }
    }
    public class txtrecord {
        public const int STR_SIZE = 0x40;

        public byte[] str;
        public uint recid;

        public txtrecord(BinaryReader br) {
            str = br.ReadBytes(STR_SIZE);
            recid = br.ReadUInt32();
        }

        public override string ToString() {
            return string.Format("str {0} {1}", Encoding.Unicode.GetString(str).Split('\0')[0], recid);
        }
    }



    public class rec_header {
        public uint mystery0; //seems to be 0x1 and becomes 0x0 when the record is deleted?
        public uint mystery1; //seems to always be 0x1
        public uint recid;
        public uint opcount;
        public uint mystery4; //always seems to be 0x20
        public uint mystery5;
        public uint mystery6;
        public uint datalen;

        public rec_header(BinaryReader br) {
            mystery0 = br.ReadUInt32();
            mystery1 = br.ReadUInt32();
            recid = br.ReadUInt32();
            opcount = br.ReadUInt32();
            mystery4 = br.ReadUInt32();
            mystery5 = br.ReadUInt32();
            mystery6 = br.ReadUInt32();
            datalen = br.ReadUInt32();
        }
    }
    public class wab_srec {
        public byte[] bin;
        public MemoryStream si;
        public int opcode;
        public int acnt;

        public bool IsArray { get { return 0 != (opcode & 0x1000); } }

        public const int MT_SINT16 = 2;
        public const int MT_SINT32 = 3;
        public const int MT_BOOL = 0xB;
        public const int MT_EMBEDDED = 0xD;
        public const int MT_STRING = 0x1E;
        public const int MT_UNICODE = 0x1F;
        public const int MT_SYSTIME = 0x40;
        public const int MT_OLE_GRID = 0x48;
        public const int MT_BINARY = 0x102;

        public const int MT_BINARY_ARRAY = 0x1102;
        public const int MT_STRING_ARRAY = 0x101E;
        public const int MT_UNICODE_ARRAY = 0x101F;

        public String ElemType {
            get {
                switch (opcode & 0xFFF) {
                    case MT_SINT16: return "int16";
                    case MT_SINT32: return "int32";
                    case MT_BOOL: return "bool";
                    case MT_EMBEDDED: return "ole";
                    case MT_STRING: return "str";
                    case MT_UNICODE: return "ustr";
                    case MT_SYSTIME: return "ft";
                    case MT_OLE_GRID: return "uuid";
                    case MT_BINARY: return "bin";
                }
                return String.Format("{0:X3}", opcode & 0xFFF3);
            }
        }

        public override string ToString() {
            return String.Format("<{0:X4}> {1}{2} = {3}"
                , (opcode >> 16) & 0xFFFF
                , ElemType
                , (IsArray ? "[]" : "")
                , Unnest(Render)
                );
        }

        private object Unnest(object Render) {
            if (Render is Array) {
                String s = "";
                foreach (Object o in ((Array)Render)) {
                    if (s.Length != 0)
                        s += ",";
                    s += "" + Unnest(o);
                }
                return "{" + s + "}";
            }
            return Render;
        }

        public object Render {
            get {
                if (IsArray) {
                    Object[] al = new Object[acnt];
                    int off = 0;
                    for (int y = 0; y < acnt; y++) {
                        int len = BitConverter.ToInt32(bin, off);
                        off += 4;
                        al[y] = GetObj(bin, off, len);
                        off += len;
                    }
                    return al;
                }
                else {
                    return GetObj(bin, 0, bin.Length);
                }
            }
        }

        Object GetObj(byte[] bin, int off, int len) {
            switch (opcode & 0xFFF) {
                case MT_SINT16: //Signed 16bit value
                    return BitConverter.ToInt16(bin, off);
                case MT_SINT32: //Signed 32bit value
                    return BitConverter.ToInt32(bin, off);
                case MT_BOOL: //Boolean (non-zero = true)
                    return bin[off] != 0;
                case MT_EMBEDDED: //Embedded Object
                case MT_BINARY: //Binary data
                    {
                        byte[] res = new byte[len];
                        Array.Copy(bin, off, res, 0, len);
                        return res;
                    }
                case MT_STRING: //Null terminated String
                    return Encoding.ASCII.GetString(bin, off, len).Split('\0')[0];
                case MT_UNICODE: //Unicode string
                    return Encoding.Unicode.GetString(bin, off, len).Split('\0')[0];
                case MT_SYSTIME: //Systime - Filetime structure
                    return DateTime.FromFileTime(BitConverter.ToInt64(bin, off));
                case MT_OLE_GRID: //OLE Guid
                    {
                        byte[] res = new byte[16];
                        Array.Copy(bin, off, res, 0, 16);
                        return new Guid(res);
                    }
                default:
                    throw new NotSupportedException(string.Format("{0:X8}", opcode));
            }
        }
    }
    public class wab_record {
        public rec_header head;
        public int[] oplist;
        public List<wab_srec> srecs = new List<wab_srec>();

        public wab_record(BinaryReader br) {
            head = new rec_header(br);
            oplist = new int[head.opcount];
            for (int x = 0; x < oplist.Length; x++) oplist[x] = br.ReadInt32();

            for (int x = 0; x < oplist.Length; x++) {
                int opcode = oplist[x];
                //Console.WriteLine("@{0:X8}", br.BaseStream.Position);
                int p0 = br.ReadInt32();
                //Console.WriteLine(" {0:X8} {1:X8} ", p0, opcode);
                Debug.Assert(p0 == opcode);
                bool isArray = 0 != (opcode & 0x1000);
                wab_srec srec = new wab_srec();
                srec.opcode = opcode;
                if (isArray) {
                    srec.acnt = br.ReadInt32();
                }
                else {
                    srec.acnt = -1;
                }
                int len = br.ReadInt32();
                byte[] bin = srec.bin = br.ReadBytes(len);
                srec.si = new MemoryStream(bin, false);
                srecs.Add(srec);
            }
        }
    }
}
