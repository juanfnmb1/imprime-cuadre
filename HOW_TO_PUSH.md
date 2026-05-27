# How to Push

Normal push:

```
git push
```

If it fails with a permission / 403 error, run this instead:

```
git -c credential.helper= -c credential.https://github.com.helper='!gh auth git-credential' push
```

That's it.
