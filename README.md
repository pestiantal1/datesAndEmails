# datesAndEmails

## FONTOS / Használat

Itt a "MUNKALAP_NEVE" legyen az a munkalap, ahol módosításokat szeretnénk végezni.

```js
const sheet = SpreadsheetApp.getActive().getSheetByName("MUNKALAP_NEVE");
```


Ha a táblázat tartlamaz első sorba címsor, i legyen 2. Ha pedig nem marad i = 1.

```js
  for (let i = 1; i <= lastRow; i++) {
    //...
  }
```

### Hibakezelés

Nincs implementálva.