Be careful with the COM Cleanup extensions, they use text templates and reflection on a particular version of the office DLL's. So you must use the version you generate against or newer otherwise there will be types missing from the office dll.

See http://jake.ginnivan.net for more information


Requirements:

To build:
http://www.microsoft.com/downloads/en/details.aspx?FamilyID=21307C23-F0FF-4EF2-A0A4-DCA54DDB1E21
http://www.microsoft.com/downloads/en/details.aspx?FamilyID=0def949d-2933-49c3-ac50-e884e0ff08a7&displaylang=en