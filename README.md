ConvWab
=======

Convert WindowsAddressBook into XML format

ConvWab is based on works of libwab/wabread distributed at filewut.com

Output XML sample
-----------------

ConvWab is a command line tool built on top of Microsoft .NET Framework 2.0 (C#)
```
ConvWab <in.wab> <out.xml>
```

Here is output:
```
<?xml version="1.0" encoding="utf-8"?>
<wab>
  <table i="0" t="index">
    <record i="0">
      <field t="ft" idHex="3008" opcHex="30080040" isArray="false" id="12296" opc="805830720" ldid="" prop="PR_LAST_MODIFICATION_TIME_STR">
        <item i="0">2014/09/08 16:38:30</item>
      </field>
      <field t="ustr" idHex="8004" opcHex="8004001F" isArray="false" id="32772" opc="-2147221473" ldid="" prop="PR_MAB_MYSTERY_PROFILE_ID1">
        <item i="0">{DE55962D-07E7-4E54-8A7C-6450C9CB7A16}</item>
      </field>
      <field t="bin" idHex="8005" opcHex="80051102" isArray="true" id="32773" opc="-2147151614" ldid="" prop="">
        <item i="0"></item>
      </field>
      <field t="ustr" idHex="800D" opcHex="800D001F" isArray="false" id="32781" opc="-2146631649" ldid="" prop="PR_MAB_MYSTERY_PROFILE_ID2">
        <item i="0">{DE55962D-07E7-4E54-8A7C-6450C9CB7A16}</item>
      </field>
      <field t="int32" idHex="800C" opcHex="800C0003" isArray="false" id="32780" opc="-2146697213" ldid="" prop="PR_MAB_MYSTERY_ALWAYS_0_NEXT">
        <item i="0">0</item>
      </field>
      <field t="bin" idHex="6600" opcHex="66001102" isArray="true" id="26112" opc="1711280386" ldid="" prop="">
        <item i="0"></item>
        <item i="1">03 00 00 00</item>
      </field>
      <field t="int32" idHex="0FFE" opcHex="0FFE0003" isArray="false" id="4094" opc="268304387" ldid="" prop="PR_MAB_MYSTERY_ALWAYS_6">
        <item i="0">6</item>
      </field>
      <field t="ustr" idHex="3001" opcHex="3001001F" isArray="false" id="12289" opc="805371935" ldid="dc" prop="PR_DISPLAY_NAME">
        <item i="0">メイン ユーザー の連絡先</item>
      </field>
      <field t="bin" idHex="0FFF" opcHex="0FFF0102" isArray="false" id="4095" opc="268370178" ldid="" prop="PR_MAB_ENTRY_ID">
        <item i="0">02 00 00 00</item>
      </field>
    </record>
    <record i="1">
      <field t="bin" idHex="800B" opcHex="800B1102" isArray="true" id="32779" opc="-2146758398" ldid="" prop="">
        <item i="0">00 00 00 00 c0 91 ad d3 51 9d cf 11 a4 a9 00 aa 00 47 fa a4 07 04 00 00 00 02 00 00 00</item>
      </field>
      <field t="ft" idHex="3008" opcHex="30080040" isArray="false" id="12296" opc="805830720" ldid="" prop="PR_LAST_MODIFICATION_TIME_STR">
        <item i="0">2014/09/08 16:39:16</item>
      </field>
      <field t="int32" idHex="3A71" opcHex="3A710003" isArray="false" id="14961" opc="980484099" ldid="" prop="PR_MAB_MYSTERY_ALWAYS_1048576">
        <item i="0">0</item>
      </field>
      <field t="int32" idHex="3A55" opcHex="3A550003" isArray="false" id="14933" opc="978649091" ldid="" prop="PR_MAB_MYSTERY_ALWAYS_0_FIRST">
        <item i="0">0</item>
      </field>
      <field t="ustr" idHex="3002" opcHex="3002001F" isArray="false" id="12290" opc="805437471" ldid="" prop="PR_MAB_SEND_METHOD">
        <item i="0">SMTP</item>
      </field>
      <field t="ustr" idHex="3003" opcHex="3003001F" isArray="false" id="12291" opc="805503007" ldid="mail" prop="PR_MAB_ADDRESS_STR">
        <item i="0">ku@digitaldolphins.jp</item>
      </field>
      <field t="ustr" idHex="3A54" opcHex="3A54101F" isArray="true" id="14932" opc="978587679" ldid="" prop="PR_MAB_ALTERNATE_EMAIL_TYPES">
        <item i="0">SMTP</item>
      </field>
      <field t="ustr" idHex="3A56" opcHex="3A56101F" isArray="true" id="14934" opc="978718751" ldid="mail" prop="PR_MAB_ALTERNATE_EMAILS">
        <item i="0">ku@digitaldolphins.jp</item>
      </field>
      <field t="ustr" idHex="3A11" opcHex="3A11001F" isArray="false" id="14865" opc="974192671" ldid="sn" prop="PR_SURNAME_STR">
        <item i="0">U</item>
      </field>
      <field t="ustr" idHex="3A06" opcHex="3A06001F" isArray="false" id="14854" opc="973471775" ldid="givenName" prop="PR_GIVEN_NAME_STR">
        <item i="0">K</item>
      </field>
      <field t="ustr" idHex="3001" opcHex="3001001F" isArray="false" id="12289" opc="805371935" ldid="dc" prop="PR_DISPLAY_NAME">
        <item i="0">U K</item>
      </field>
      <field t="ustr" idHex="3A51" opcHex="3A51001F" isArray="false" id="14929" opc="978386975" ldid="businessUrl" prop="PR_MAB_BUSINESS_URL2">
        <item i="0">http://www.digitaldolphins.jp/</item>
      </field>
      <field t="ustr" idHex="3A26" opcHex="3A26001F" isArray="false" id="14886" opc="975568927" ldid="businessCountry" prop="PR_MAB_BUSINESS_COUNTRY">
        <item i="0">Japan</item>
      </field>
      <field t="ustr" idHex="3A28" opcHex="3A28001F" isArray="false" id="14888" opc="975699999" ldid="businessSt" prop="PR_MAB_BUSINESS_PROVINCE">
        <item i="0">Osaka pref.</item>
      </field>
      <field t="ustr" idHex="3A2A" opcHex="3A2A001F" isArray="false" id="14890" opc="975831071" ldid="businessPostal" prop="PR_MAB_BUSINESS_POSTAL_CODE">
        <item i="0">544-0013</item>
      </field>
      <field t="ustr" idHex="3A27" opcHex="3A27001F" isArray="false" id="14887" opc="975634463" ldid="businessLocality" prop="PR_MAB_BUSINESS_CITY">
        <item i="0">Osaka city</item>
      </field>
      <field t="bin" idHex="0FFF" opcHex="0FFF0102" isArray="false" id="4095" opc="268370178" ldid="" prop="PR_MAB_ENTRY_ID">
        <item i="0">03 00 00 00</item>
      </field>
      <field t="int32" idHex="0FFE" opcHex="0FFE0003" isArray="false" id="4094" opc="268304387" ldid="" prop="PR_MAB_MYSTERY_ALWAYS_6">
        <item i="0">6</item>
      </field>
    </record>
  </table>
  <table i="1" t="text">
    <text id="3">U K</text>
    <text id="2">メイン ユーザー の連絡先</text>
  </table>
  <table i="2" t="text">
    <text id="3">U</text>
  </table>
  <table i="3" t="text">
    <text id="3">K</text>
  </table>
  <table i="4" t="text">
    <text id="3">ku@digitaldolphins.jp</text>
  </table>
  <table i="5" t="text" />
</wab>
```

Compare with wabread's output
-----------------------------

```
version: 1
#
dn:: Y24944Oh44Kk44OzIOODpuODvOOCtuODvCDjga7pgKPntaHlhYg=
cn:: 44Oh44Kk44OzIOODpuODvOOCtuODvCDjga7pgKPntaHlhYg=

# U K
dn: cn=U K
cn: U K
mail: ku@digitaldolphins.jp
mail: ku@digitaldolphins.jp
sn: U
givenName: K
businessUrl: http://www.digitaldolphins.jp/
businessCountry: Japan
businessSt: Osaka pref.
businessPostal: 544-0013
businessLocality: Osaka city
```
