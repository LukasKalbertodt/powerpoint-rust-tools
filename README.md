PowerPoint Rust Tools
=====================

PowerPoint Add-In that offers some tools to work with Rust code on slides. Currently it's only a "syntax highlight the Rust code in this text box"-tool.

![scrot](https://github.com/LukasKalbertodt/powerpoint-rust-tools/blob/master/screenshot.png)

***Beware***: this Code is **horribly** written! I don't know anything about C# nor MS Office development. The code suits *my* needs and contains many hard-coded values. In fact, I'm pretty sure this is the worst code I wrote in the last few years ...

Feel free to try to improve it, though :wink: I'd be happy to accept a PR.

## Installation

You need `pygmentize` to use this addin. First just install Python 3.x from the website. Then:

```
$ pip install Pygments
$ pip install pygments-style-solarized   # you want solarized!!
```

And make sure to put the folder of `pygmentize.exe` in your `%PATH%`! It's usually `c:\users\<user>\appdata\local\programs\python\python37-32\Scripts` or something like that.

Next, download the `installer.zip` from this repository. Extract it and execute the `.exe` inside. Always execute executable from strangers from the internet!
