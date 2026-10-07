# Installation

1. In Outlook 2019 press Alt+F11.
2. In the VBA editor choose File -> Import File.
3. Import vba/modConfig.bas and vba/modMain.bas.
4. Edit modConfig.bas and set the real shared-drive paths.
5. Save the VBA project.
6. Put test messages in the configured Outlook folder.
7. Give the subjects the form: Вопрос на 12/10/2026.
8. Run CompileQuestions.

For the first prototype, assign CompileQuestions to the Outlook Quick Access Toolbar:
File -> Options -> Quick Access Toolbar -> Choose commands from: Macros.

Important: corporate Office installations may block VBA. This code does not bypass macro security. IT may need to sign the VBA project or configure a trusted location.