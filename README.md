# 40FINGERS StyleHelper-Sko for #dnncms

About 40FINGERS StyleHelper Skin Object for DNN

The basic function of the 40FINGERS Style Helper Theme object is to allows you to manipulate the CSS, Javascript links and meta tags and content of you DNN page from the Theme.

The StyleHelper allows you to do this only when certain conditions are met.
Some examples of the available conditions are: browser type / version, username, role, query string etc.
These filters can be positive (include) or negative (exclude).

And it can do much more, check out the [documentation & release download](https://www.40fingers.net/Products/DNN-Stylehelper)

Demo Themes: https://github.com/40fingers/StyleHelper-Sko-DEMO-Theme

### Note: I sometimes get questions: Why is all the code all in one file?

When I first created the Style Helper, some clients did not want to install any third party extensions.
So a dynamically compiled Skin Object was a good way around that limitation as you can also load the SKO from the Skin folder and include it in the Skin package.
Although it works fine, now that the project grew a lot it's not ideal to maintain and the next Major version will be compiled (and in c#).
