# <img src="https://webapps1.chicago.gov/cdn/FontAwesome-5.13.0/svgs/solid/pen-nib.svg" width="32" height="32"> CreateSetOutlookSignatureDefault

This script is designed to automatically generate and set Outlook Classic signatures for users. It is context-aware, meaning the signature content changes based on AD user information and the time of year (e.g., a standard signature vs. a Christmas/Holiday signature).

## Core Logic

The core logic revolves around:

* [Signature Themes](https://github.com/eduardomozart/CreateSetOutlookSignatureDefault/wiki/Signature-themes): The file [signatures.csv](signatures.csv) acts as the controller. It determines which signature folder to use based on the current date.
* Template Files (``.txt``, ``.htm``): These contain the layout of the signature and include placeholders for user information and dynamic dates.
* [Variable Replacement](https://github.com/eduardomozart/CreateSetOutlookSignatureDefault/wiki/Variable-replacement): A processing step that reads the templates and replaces placeholders like ``%%FirstName%%`` or ``%%Year%%`` with actual data.

## Functionality

- [x] Photo: Extract LDAP photo to JPG file, similar to [syncLdapPhoto](https://github.com/glpi-project/glpi/blob/815e2a9690481ce89b03d06cd35a28c36f6d24b6/src/User.php#L1724-L1801) (GLPI)
- [ ] ~Extract user group information to apply signatures based on it.~ (Use ``%%Title%%`` variable instead)
- [ ] New Outlook support.
- [ ] Roaming signatures support (Microsoft 365/Outlook.com).
- [ ] ~Add support for embedded HTML signatures to Thunderbird.~
- [ ] Conditional variable fields in Templates (If...Else...End If).
- [x] RegEx replace for variable fields (e.g., strip all non-digits from ``%%Phone%%`` variable).
- [ ] Convert it to PowerShell when the VBScript Feature on Demand (FOD) disables it by default (~2027).
- [ ] SVG signature with user variable support, including conversion to PNG.
- [ ] Exchange: Allow exclusion/inclusion list of users/domains with RegEx support.

## Documentation

Refer to [Wiki](https://github.com/eduardomozart/CreateSetOutlookSignatureDefault/wiki) for more information.

## Useful links

<WRAP center round tip 60%>
Você encontra modelos gratuitos de assinaturas de e-mail em https://www.mail-signatures.com/
</WRAP>
