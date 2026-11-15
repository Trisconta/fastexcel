# History of fastexcel

## At gist

`fastexcel` Python3 script was born on gist the first time on May 2020, to allow myself comprehend how openpyxl worked.

It grew a bit as a standalone script, but not significantly more.

I realized I used it for scripting quite a lot, and I kept it simple enough, aiming simplicity of use and coding.

## Final words... at Trisconta

Now this got moved into [Trisconta]() repository, for its final place.
I challenge anyone to fork it and make your own customizations!

Here are the commands I have played to move it to Trisconta `fastexcel` own repo:

```bash
git clone git@github.com:Trisconta/fastexcel.git
cd fastexcel
git remote add other git@gist.github.com:fa9188a3fe219ef3e72371a00c633192.git
git fetch other
git rebase other/master
git push --force
```

(c)2026  Henrique Moreira `<henrique@declaratived.com>`
