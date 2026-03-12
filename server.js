"use strict";
const express = require("express");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle,
  WidthType, ShadingType, VerticalAlign, PageNumber,
  TabStopType, TabStopPosition, PageOrientation, SectionType
} = require("docx");

const app  = express();
app.use(express.json({ limit: "10mb" }));

// ── API Key Authentication ────────────────────────────────────────────────
const API_KEY = process.env.IPMIND_API_KEY;

function requireApiKey(req, res, next) {
  // Health check is public
  if (req.path === "/") return next();
  // If no key configured on server, block all requests for safety
  if (!API_KEY) {
    return res.status(503).json({ error: "Service misconfigured: IPMIND_API_KEY is not set." });
  }
  const provided = req.headers["x-api-key"];
  if (!provided) {
    return res.status(401).json({ error: "Missing x-api-key header." });
  }
  if (provided !== API_KEY) {
    return res.status(403).json({ error: "Invalid API key." });
  }
  next();
}

app.use(requireApiKey);

// ── Embedded logo (base64 PNG, navy background) ──────────────────────────
const LOGO_B64 = "iVBORw0KGgoAAAANSUhEUgAAAUAAAABpCAYAAABLRPrgAAAABmJLR0QA/wD/AP+gvaeTAAAY4klEQVR4nO2debQdVbGHf8UkyEwMiRAiCkQiyDwKKAJRkEFRRuGpCD4BEVBQUViCPBFkkMElLMUBUZmePomggEBkkEEgAQXDPIYkkBBISEiAJPd7f+x9vCcnfXo6Q5+bW99aWRfO6V27uvt09d61a1dJjuM4juM4juM4juM4juM4juM4juM4juM4jjPgsaoVKAvwTknvkTRM0rskrSxp2bpD5kh6XdJ0SS9Kmmpmfd3W03Gc3mVAGEBgbUk7SNpa0uaSRktaq6CYtyQ9KelhSeMl/UPS/Wb2VhtVdRxnANGTBhBYWtJOkj4paQ9J7084bI6k5yVNkfSKpNmS5kuaJ2kVSctIWkPSUEkjFQzm0g0y5kq6Q9L1kq41s8ntPhfHcXqXnjKAwAckHSHpYEnD676aIelOSfdKmiDpETObWlD2cpJGSdpU0lYKI8ot1G8U+2Ifl0m6xszmlj4Rx3GcPAAG7A2MY1H+BZwGbAks1aG+VwcOAK4AXq/r+zXgHGBEJ/p1HGeQEw3fvsA/GwzP+cAHK9BnBeDgBkP8FnCJG0LHcdoGsB1wT52heQo4Mq7sVg6wMfAr4O2o31zgDGClqnVzHGeAAqwGXAr0RcPyPHAYsEzVuiUBrBsN4cKo7yTgU1Xr5TjOAAPYHXgxGpI3gJOBFarWKw/ApsAddSPWK4E1qtbLcZweB1gu+vVqo75bgPdWrVdRos/y8OinBHgB2KFqvRzH6VGA4cBd0WDMA44Geir0pijACODWeE5vA8dWrZPjOD1GnDZOiobiCWCTqnVqF8BShBCdmm/wp73qx3Qcp8sAHwFmRePwV2D1qnXqBMA+wOx4nn8aKD5Nx3E6BDAmho0AXAYsm91q4AJsAUyN53srPRLK4zhOlwF2qTN+5w90f19egPVjSE9tkWf5qnVyHKeLAFvRv53sR1Xr022A99YZwf8jJHRwHGdJB1inbhp4aadHfsB6hOQGRdqMAFbplE6xj1HAtMH6EnCcbgMMJeQLaPZv3U4rsAIwIT7013V65ENIUgBwX15DC+wXV2ynA0XzCBbVb5s6N8Bhneyrod/VCLtqfgCcCCSlEHMGKMBawFHAWcAxhByZgx7CFto0Luu0Ar+MHT0MrNzRzkJ/tbjCBcCKOducW3dBPtoFHQ8iBH7PAzbrQn97Aa823PgF8WEZFH7YJZn4kM9tuL/zgKOr1q1qKjWA8UGHEPKyQcc6WrTP7YBrKTC6Irw9rwDOpEOptRL6PD9em0fp4MowYRX6zZQfwDc61bfTeQgZk/qa3Ns+YN+qdaySygxgNCozYief7UgnAxjCFsAH4vW5qIP9XJvxA5iFxycOWIBHMu7vv6vWsUqqNIC1B++ajnSwBACMJkxVFtKhfcN1L6E0tu9E305nAYbkuLcA76pa16po1QCWmg4CeyvU65gh6ZgyMqKckYTYwdw3MP4odqHA6g7wTmBHYBPyL5wsRZhefgh4R96+6jGzRyWdrnCdL6Ez2+XyLDp5SM7AJO998/vbLQhTuyejdT2iBTln0u/bmAscmqPNgcCc2KYPOC9Hm83oj8+DEKic6pMjpMq/u67NE8CoIudXJ2tZ4N9RzlfKyMiQf1PGG3AeXViccjoD8EzG/X2mah2rhG5PgQlL8BBCX8qOID+aoOhcYFhKmzXoN3717J7R1wMJbU7PaHNRQpuby5xrlPfxKONl2pxVGtiZ/qQMSZzVzv6c7kJIwZbG4VXrWCV00wASYv6mRMEfa0HOyU2U3SOlTZLRhBRjFvVNWkG7JUO/8Qlt5tJCSAlh5AnwnbIyUmQfQahf0sjvWML3Yg8GCOUYGn/HfcAPqtatauiyAfxKFHpni3IOSlC0j5RQGuA9CT8CgM9n9DUpoc1PM9pck9Dm4bLnG2V+KMqZTgfCYoD3Ad8jZKv+MbBLu/twqgPYHPghcDVwNrBl1Tr1Al0zgIRFgaei0KYjtZyylgFub1D0whztzmloczcZW+KA/YH5dW2mAiMz2oxm0dXVecDHi55ngtzboryjWpXlOE53DeCeUeBE2rC7IBrBQwmjllQ/XkO7MbHN58g5vSOs/p4MfJWcK87Au4HjgZNo07YyQv5AgEfaIc9xBjvdNIBjo8CvtkXgIIQwin4uXkePzXOcFmnVAOaKSwOGStpD0puSftu62oMTM+uLN+RUSZ+XdE+1GpUH+ICkTSUNl7SWpLmSpkt6QdIdZjazQvUKQajwt6ukEZLWljRP0jRJUyT9w8xezCFjWUnbShoVZawm6Q1Jr0l6SNL9ZjanIyfQAYANJW0jaU1JwyTNVziX5yXdY2aTOtTvaElbSRoq6d0Kv6tpkp6VdJuZze1Ev1lKfTla0z90vfMlDGCDeC2n0YbA6DitTyN1pw7BFZHGr+uOXQE4AXg6o818gr/zMLqcFxH4doZuZ9UduzXwZ5JX0Ou5F/hMk/6GEBYnsnbkvAVcBWxV8HwuzJB7Rkb70Rntn607diXgu4S41ywmErLTtLzNklA87Rz61xiaMRf4I7BrXdvOT4GBG6Kwg1sW5gh4MF7Pndsgq9MG8Np43NZkB+Um8SiwZ6vnmRdyGMB4zucTMuYUYSx1OSUJfvFXCspYQEhZliuGli4ZQELKuCkZxyYxEfhgyXu1PCHEJym+N4urgJXp9FY4QsjGzgpD4BvKnGirEH6wxwH/Iowu5hFWkT+Zo+1wQlhIrSj7DEJ83Ibd0L0J18W/n6hQh7ysQtjH/DdJZeo5byjpOuBUeiM119KSfi3peBXfQraPwrksBxwiaaykISX6/7akjiXIKApwjKSrFaacRRkt6e9FjSBh08M4Sd+RlCutXQMHSrpZUmdr8AAfi4bj7x3tqHn/y5G+3evMlLYbApObtJsDjOnmudTptWPUYXwbZHV6BPgE/S+PVvllq+eb43pkjQBntuE8riI9BVleMlNZ0fkRYNFRcDOeJed0mJCd/bl0cbnJGoFflqZLnmH4TvHv3/KcXFHIDgr+nqS0XScnkTASJPierlFw0CexoqSrgKZvcGBpCqbez8k/FJy7m9LhNP1tYAMFp347OAw4vk2yyrJqG2QcKKlUgowGmr68u0i7fLTrSsqMbyUkFvmDpPe0qd+iI/BFyGMAt45/27piSfA5vCjpDeAhEpzD8WLlSSDw9YTPxkjKGpavIekLCf0uA5wvaZak2YTdFavl0CMXZjZf0n0KP75t2iV3gHAOsG3VStSxUGG1d1qLcpD0kqSp8b/z8H5gkxb77QRIekXSJElvFWiXJ8D/PPXQbz6PAdwi/i08XSPEvZ1IWFlavu7zjSRdof6RxaaSxrJ4ooDRkvJkMtmWxf1LeS/ydgmfnaDgI1pR0nKSDpJ0Qf0BhCQEFwIb5+ynkdr13Lxk+6qYIukHCn7h9eO/nSWdEb/LYpl4bNW8qvDADjGztc1smKR1FHxzeQ2YJM1WeAGvaWbvNrO1FEKDzlAwrll0vERDAZ5QGBAMNbOhZjZS4fkbo3wDoPWBpn5iYD1JX86py18lfVHSRgqhOO+X9ClJv5G0IKeM1gDWjPPol0q2/0zdXPyYus+/2WS+vmtD+20y5vc15tMQbgF8P2fbsQl635Nw3My675ei35f0QMlr8/nY/rIy7evkdNoHWM95pLgsCHkXz8spa8dWzjtFhywfIIQs2aNTZJyU8xzmAE1fYMB/55CRmq2HzvsAa9xASo0dQlq3G3PIaerXpL9+UBrTgN0yzmlDwoJoHi5Lk5U1AqzlwHsi47hmTFJ4C6IQQFljVpPjGz9/XGH1OYuJZtb4ts273SwpycHraZ+ZWZ9CwK+06HkVoXZNu1JLpQ18y8xOSAtENbO5ZnaCpJNyyMs7EugEp8Rktc04V+G3m8X3zezBlO8vlZSVsr4lH1abmCrpADN7o9kB0W1zlLJHtYlbTQk5KQ/JaPuqpB3MLDVbk5k9JmlHhQDzlsgygDVH5XNlhJvZfQp+uC3N7Lq6r66W1Bhdf5ukCQ3tZ0n6fY6ufp7w2fWSXs5oN18hJKKR87X4NKixzu9OClODsvVQagGo7XIGd5KxZnZ2gePPVnbI1BiqCYuZJ+lXaQeY2QKFKVgafQoGLk0Oyr4OXSnSlcEFZjY76yAze1bZrrBmvvKdFdxJaXzJzJ7M0iPq8rqk/RV2p5Um6+LX4oIml+3AzB5tfEvGbVLbSvqJpFsk/Y+kveLIqpGvqd9YJHGTpEsS+p0j6TBJb6e0/UbSBTezGyXtJul/FQzpoWZ2QcMxs8zsFjMr4iSu52UFAzycLu+WKEifpJOLNIgP/ikZhw1T8P12mwk5t6RljQAnmtmMHHJKPztd5LrsQ/7D0xnfN7MpWflDJ0j6YwE9ZGZPSWoptCrLANaGs62ukC2GmU0xs2PMbIyZfbfZ8NvMXpa0vaQrteh0eLaCM36f+MZOanuDpF0k3d/w1dOS9jezpim4zGycmR1gZnub2e/yn1k+opGYobAS3LYV5g4w3swKVx4zswmSHss4rOM1kxPI685Je3FK+abIUj4XTpXMV/Z9qidzpNiErHt9RXwmipI0g8tN1l7UWoxakk+sa0Qj+FngSEkfUFian5hn9GVmd0naBlhH0khJ082srE+z3cxUWDFcVcEY9iKtxH/errATpBlDW5Bdlmb+56rkVM2MkoanKFn3+vaScscrGOVSdW+yDGAtdGVeGeHtJs777y3ZdpLyv7W7Re26Lp96VLW8kH1IU7Kmf2u2ILssZV0WjSS5awYiLfnQCtC03k8kM+NOEma2EJis9BdtU7KmwLXvl5Sb3WvUpu6dKJfZLlpJ4ZQ1c+j1XTBO+8gqBtbJ31lTsgxgzfD1wkrVkkjN8PWyn2iNFtpmhXi80oJsZ2Dxasb3rfzOSheGzxp51IbHLef8ageEYugbKRiM8TlX4URIPbS5QsLL6ZIeMLMsJ3c3qAUV94SLoQml6iFH1sv4fnoLsp2BxTSluzzWVwl3CyEwv/Re9ayRXW1oWelUhVDx7GaFcJjrFUJfXgJ+TkadXWAvSU9JekDStZLukjQFOLrDauehtjG/lx3qpTLmxBi/rMp0WXGazpJD1r1O3f2RwofVQmKKLANYm6K03Vkdt7NcQ0gOeglNihXFUd/dWvwCLSPpcEk3U7fPuKHtQQo52xr3Jw6R9BPSawrvD9xKqDx3LCWLwKfIt6jHfIXV4F5lPQoUrapjX4UV7jRKLWg5A5K7M74/jHLZpY/JPqQ5WQ/11Pi39BCTUA93t4bPhkv6u0Ik92aSjpR0I8kp4i9U+grSdgqJCxr7XU3SxUo/x5NJSOQI7K+QSmsXhRjEC9UQ2EuoGncg5dNZDZO0rKSpXQpDaIVzKFDLmLCn9PsZhz0edxY4g4OsnTXDlW8L5X+IL+aWkgpnGcDaPtcymYAFfFhhynkzcEDdV/tpcQf5lupPvVVrP0RSnnTqScXR95G0eka7pST9V8LnSftUj2z4/zslXaWwra8M74t/nyvZvptsLOlyQnqyVOIxlytk8kmjkuziTmXcq+yFkFPirC0TYFOFzDAtbafMMoC1gOGyjvD65JP1ux2a7QlsfMDWV76EjRskbCfLW8s3KX4oaUr9H93idLg28iu7i6PVRBPd5jOSbgO2aHZA/O42SZ/OkDVf0o/bp5rT68TdWln3fClJVwDn0iT/JiF70ZEKM8jSq781UleBzWwaMF3SmsCwuCOjCNcr5PR6pxbdszdWYf9v/bTqaYVMyfXkXR19KyEbTN62SdlNrpK0Q8Jnkv5T3nJ3SXspvIXKUJt6D6Qi6dtJuh+4R2EP93Px83UVykruoHxv5MvN7JlOKOj0NBdIOk7pgwZTyMd5BHCDgk2YphAmM1phZjeiXQrlCcCdIOnjClPUvxQRHn1bi2XeMLOngb0V0g6NUki2eIyZNRqtRxWGzVkxQkn1Su7KqWZS24sVRq9HKYQAXSnpm/UHxL2uExZvmpst49+WU/p0maUUDF3jCyIv0yWd1jZtnAGDmc0ETlZIgpLFqgqJiHNNicuSZ2Xzvvj3Q+3sOCYb2MLMVooJER5POGa+QgrtNJD0w4TPb1P2KuNUJWymNrM+MzvDzEaY2ZCYtKFtBZkJdUa2VpgKlkqoOkBZKOngPIXGnSUTM7tYwUfcE+QxgHfEvzt3UI80fqi66WcDfZJONLNxjV/E0eeBkprlF5shad88edA6wLYK0/8JaUkoe4jGbDplWCDpSDO7tQ2ynIHNkcpeFc5L04xOechjAO9S8KdtC2Stqrad6Nv7bPw3TiGgcrJCZamdzKwxUWl92xckbSXpVIWp5qsKWaYvlLSJmTX6HLtFLTdaaubbHmJ/FUuZ1MhrkvYws6TEtc4gI7q69lTrC2HnKiTfLU2mAYzK/k3BX1hJIW8zw8yuNLNdzWx4nJruZ2ZZwZUys9fN7HQz2zxOZzc0s+PNLE8Bn06xd/xbyKdaIVMUXCC/U7GCQQsUsnVvkpXm3BlcmNkCMztW0h6S0soKJDFD0hfM7Bsqn59QUv4sJNcqGL/9FB4CpyTAKIVMyFPUnp0QTypkrm5GW3ZbmNlrkg4FfiTpUIWdHusmHIrCD/ovkn5TQe7FR5V+PZJqwJSRk3f28FSGnCwf8EMZ7bOiCF7PaF90P/b4DHmFZgpmdiNwk8KgYF8FO5O086xP4VpcKekX8fcohSwyafqkum9yBRHGbWpTohJrmVlWQKPTBOB7kr4r6UIzq7pIuOLum6xsNMvFBanGtmsr7BIarpBnb4qkyf77cMoSY2xHKPyuhiq8UKdLeiyW0qhMsT/GMnPHVabEAIdQTvP5eB2bBhR3E/KVxVy2aj0dp1KAT8SH4THaUM2LUEP2OODHwKF5kg0ABhwc23yNlDqmDe0+QqhXexowMmeb9YHTgbOBtlSyBz4Zr2HhIvOdwg2g4+Qgjl6ejA/EXi3KWh54qOEhy9xRAfyioc0jZGzSB74E9NW1mQmkbpMDtgDeqGuzANiv6HkmyL0jyvtCq7LahRtAx8kJcFR8IDJXXzPkfK7Jg7ZRSpsNmrT5UkZfLye0Sa0LS/90v55WwkAE7BjlvEgIhO4J3AA6g5miOe5+peDo3h7Yo4V+m+3lS9vj1ywlV9M2hDyBSRums/YSJn0/gtam/rXcg+f0SDZqx3GKAhwdRwUPUbKgN7A9i05LAWYR0l81a7Mq8GrC6OQjGX3dmdAmNe8Ywe/XSJHi0Y3yav7TSTRJ3loVPgJ0nAIAywKPxwejMUdeETknAm9HOa8SkiNktdkdmB7bzAe+k6PNKIKvsMbVZExBgZWAP9e1GU/OxZMEWcsRFo4AkvIWVoobQMcpCLBnneHKqveZJmdVYDMKpMIG3hHb5M7DBywNjAYKpdEB1iWk7i+dDh84JV6re1uR0yncADpOCYA/xIfj91Xr0qsAGwHzCKPVzavWJwk3gI5TAkJNjBnxAflc1fr0GnGk+mC8PmdWrU8z3AA6TkmAA+IDMhtISi3fiT7HEBY2vl6gzfuAm4CfkVx4qe0AF8Vr8yA5amlUhRtAx2mBaFQAJlK+QlqR/u6N/fWRfyfIj+oe5rL1R4voeGjs63Uygq6rxg2gM5hpx2joWElbKKR4vxrYOxZA6RTXKyQUHVcgmejNko5WyCX4z04pJoUyoJIuVdjEfXhSpuseo0/Sz3Ic4zhOEsAIYHIcLfyCNuwVzuhv9aJ9ACt3eiRDWDGuhemc1sm+HMfpIYDNCcHMABdUrU+3Ad5LCHQG+HWnXwKO4/QYwM7A3GgELhosRoAQbP1CPO9r6dJCi+M4PQawG/2ZVC5f0h3owNb0J1z4Ez284us4ThcAdgJei0ZhHCl7fAcywKeBOfE8r1jSjb3jODkBNgaei8bhGXokA3I7IGytO4P+hA5nDZbpvuM4OQGG0Z+N5U1CFuee2w9bBML+4Fpi03n0YIIDx3F6BEL2mHPqRku3AetXrVdRCBmxv1y30v04sFnVejmOMwAgbGF7oW7kdBoZ6ex7BWBL4O6oex/wU2ClqvVyHGcAAawCXAwsjMbkRULNjp5cPCAURvpNnb5PALtWrZfjOAOYOKKqz9T8LPDVXhlVEfIN/paQxoo47f0WHuLiOE67APYGJtQZwpmEAOqurxgDKxKKNd1ep8+c6L8c2m19HMcZBBBq/e4B/JVF64Q8Sgg12Z6StUdy9D0UOAS4hkXLYL4EnAokFVRyHGcJoadi14BRkr4o6RAtWpltlqQ7Jd0rabykiZImmRkFZK8gaUNJm0jaWtJOkj6o/mswXyFrzGWSxnrlNsdZ8ukpA1iDECu4naRPSdpd0sZaXNc3JT2rkOJquqTXas3jsctLWl3SUEkjJa2V0NVMSeMUUmz9ycxmtPVEHMfpaXrSADYCrClpB4U8gJtJGi1pHRXTf7akxyU9LGmCpLsl/dPMFrZXW8dxBgoDwgAmEVdkR0oaLmmIpJUk1eIKF0p6XWHq/LKkyWb2ShV6Oo7jOI7jOI7jOI7jOI7jOI7jOI7jOI7jOI7jOE4H+X82HH2b5xE8bQAAAABJRU5ErkJggg==";
const logoData = Buffer.from(LOGO_B64, "base64");

// ── Palette ──────────────────────────────────────────────────────────────
const C = {
  orange:     "FF6734", navy:      "0F1F38", ink:       "1C1C2E",
  mid:        "4A4A6A", muted:     "7A7A96", rule:      "E2E2ED",
  surfaceAlt: "F4F4F0", amberText: "8A5A00", amberBg:   "FDF5E0",
  greenText:  "1A6B4A", greenBg:   "EAF5EF", white:     "FFFFFF",
};

// ── Page geometry ─────────────────────────────────────────────────────────
// A4 Portrait  content width = 11906 - 1440*2 = 9026 DXA
// A4 Landscape content width = 16838 - 1440*2 = 13958 DXA
const PG  = { W: 9026  };
const PGL = { W: 13958 };

// ── Border helpers ────────────────────────────────────────────────────────
const noBorder  = { style: BorderStyle.NONE, size: 0, color: "auto" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };
function solidBorder(color, size = 4) {
  return { style: BorderStyle.SINGLE, size, color };
}

// ── Shading ───────────────────────────────────────────────────────────────
function shade(fill) { return { fill, type: ShadingType.CLEAR, color: "auto" }; }

// ── Margins ───────────────────────────────────────────────────────────────
const CM  = { top: 80,  bottom: 80,  left: 120, right: 120 };
const CMW = { top: 120, bottom: 120, left: 160, right: 160 };

// ── Text helpers ──────────────────────────────────────────────────────────
function run(text, opts = {}) {
  return new TextRun({ text, font: "Arial", size: 20, color: C.ink, ...opts });
}
function para(children, opts = {}) {
  if (typeof children === "string") children = [run(children)];
  return new Paragraph({ children, spacing: { after: 80 }, ...opts });
}
function emptyPara() { return new Paragraph({ children: [run("")], spacing: { after: 0 } }); }

// ── Section heading ───────────────────────────────────────────────────────
function sectionHeading(text) {
  return [new Paragraph({
    children: [new TextRun({ text, font: "Georgia", size: 32, bold: true, color: C.navy })],
    spacing: { before: 480, after: 160 },
    border: { bottom: solidBorder(C.rule, 4) },
  })];
}

// ── Sub-heading ───────────────────────────────────────────────────────────
function subHeading(text) {
  return new Paragraph({
    children: [new TextRun({ text: text.toUpperCase(), font: "Arial", size: 16,
      bold: true, color: C.navy, characterSpacing: 40 })],
    spacing: { before: 200, after: 80 },
  });
}

// ── Claim block (left orange border) ─────────────────────────────────────
function claimBlock(label, text) {
  return new Table({
    width: { size: PG.W, type: WidthType.DXA },
    columnWidths: [60, PG.W - 60],
    borders: noBorders,
    rows: [new TableRow({
      children: [
        new TableCell({
          borders: { top: noBorder, bottom: noBorder, left: solidBorder(C.orange, 12), right: noBorder },
          shading: shade(C.white), margins: { top: 0, bottom: 0, left: 0, right: 0 },
          width: { size: 60, type: WidthType.DXA },
          children: [emptyPara()],
        }),
        new TableCell({
          borders: noBorders, shading: shade(C.white), margins: CMW,
          width: { size: PG.W - 60, type: WidthType.DXA },
          children: [
            new Paragraph({
              children: [new TextRun({ text: label.toUpperCase(),
                font: "Arial", size: 16, bold: true, color: C.orange, characterSpacing: 40 })],
              spacing: { after: 80 },
            }),
            new Paragraph({
              children: [new TextRun({ text, font: "Arial", size: 19, color: C.ink, italics: true })],
              spacing: { after: 0 },
            }),
          ],
        }),
      ],
    })],
  });
}

// ── Summary card table ────────────────────────────────────────────────────
function summaryCardTable(cards) {
  const colW = Math.floor(PG.W / cards.length);
  return new Table({
    width: { size: PG.W, type: WidthType.DXA },
    columnWidths: cards.map(() => colW),
    borders: noBorders,
    rows: [new TableRow({
      children: cards.map(c => new TableCell({
        borders: {
          top: solidBorder(C.orange, 8), bottom: solidBorder(C.rule, 4),
          left: noBorder, right: solidBorder(C.rule, 4),
        },
        shading: shade(c.highlight ? C.amberBg : C.white),
        margins: CMW,
        width: { size: colW, type: WidthType.DXA },
        children: [
          new Paragraph({
            children: [new TextRun({ text: c.label.toUpperCase(),
              font: "Arial", size: 15, bold: true, color: C.muted, characterSpacing: 40 })],
            spacing: { after: 60 },
          }),
          new Paragraph({
            children: [new TextRun({ text: c.value, font: "Arial",
              size: c.small ? 20 : 28, bold: true,
              color: c.highlight ? C.amberText : C.navy })],
            spacing: { after: 0 },
          }),
        ],
      })),
    })],
  });
}

// ── Mapping item ──────────────────────────────────────────────────────────
function mappingItem(num, featureText, conclusion, rationale) {
  return new Table({
    width: { size: PG.W, type: WidthType.DXA },
    columnWidths: [400, PG.W - 400],
    borders: noBorders,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: solidBorder(C.rule, 4), bottom: noBorder,
              left: solidBorder(C.rule, 4), right: noBorder },
            shading: shade(C.white), margins: CMW,
            width: { size: 400, type: WidthType.DXA },
            verticalAlign: VerticalAlign.TOP,
            children: [new Paragraph({
              children: [new TextRun({ text: String(num), font: "Arial",
                size: 24, bold: true, color: C.white })],
              alignment: AlignmentType.CENTER, shading: shade(C.navy), spacing: { after: 0 },
            })],
          }),
          new TableCell({
            borders: { top: solidBorder(C.rule, 4), bottom: noBorder,
              left: solidBorder(C.rule, 4), right: solidBorder(C.rule, 4) },
            shading: shade(C.white), margins: CMW,
            width: { size: PG.W - 400, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: featureText, font: "Arial",
                size: 20, bold: true, color: C.ink })],
              spacing: { after: 0 },
            })],
          }),
        ],
      }),
      new TableRow({
        children: [
          new TableCell({
            borders: { top: noBorder, bottom: solidBorder(C.rule, 4),
              left: solidBorder(C.rule, 4), right: noBorder },
            shading: shade(C.surfaceAlt), margins: CMW,
            width: { size: 400, type: WidthType.DXA },
            children: [emptyPara()],
          }),
          new TableCell({
            borders: { top: noBorder, bottom: solidBorder(C.rule, 4),
              left: solidBorder(C.rule, 4), right: solidBorder(C.rule, 4) },
            shading: shade(C.surfaceAlt), margins: CMW,
            width: { size: PG.W - 400, type: WidthType.DXA },
            children: [
              new Paragraph({
                children: [
                  new TextRun({ text: "Conclusion:  ", font: "Arial", size: 19, bold: true, color: C.ink }),
                  new TextRun({ text: conclusion, font: "Arial", size: 19, color: C.mid }),
                ],
                spacing: { after: 80 },
              }),
              new Paragraph({
                children: [
                  new TextRun({ text: "Brief Rationale:  ", font: "Arial", size: 19, bold: true, color: C.ink }),
                  new TextRun({ text: rationale, font: "Arial", size: 19, color: C.mid }),
                ],
                spacing: { after: 0 },
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

// ── Justification panel ───────────────────────────────────────────────────
function justificationPanel(text, W = PG.W) {
  return new Table({
    width: { size: W - 80, type: WidthType.DXA },
    columnWidths: [48, W - 128],
    borders: noBorders,
    rows: [new TableRow({
      children: [
        new TableCell({
          borders: { top: noBorder, bottom: noBorder, left: solidBorder(C.orange, 10), right: noBorder },
          shading: shade(C.white), width: { size: 48, type: WidthType.DXA },
          margins: { top: 0, bottom: 0, left: 0, right: 0 },
          children: [emptyPara()],
        }),
        new TableCell({
          borders: noBorders, shading: shade(C.white),
          width: { size: W - 128, type: WidthType.DXA }, margins: CM,
          children: [
            new Paragraph({
              children: [new TextRun({ text: "ESSENTIALITY JUSTIFICATION",
                font: "Arial", size: 15, bold: true, color: C.orange, characterSpacing: 40 })],
              spacing: { after: 80 },
            }),
            new Paragraph({
              children: [new TextRun({ text, font: "Arial", size: 19, color: C.mid })],
              spacing: { after: 0 },
            }),
          ],
        }),
      ],
    })],
  });
}

// ── Analysis paragraphs ───────────────────────────────────────────────────
function analysisParagraphs(interpretation, mappingDetail, differences, opinion) {
  // mappingDetail may be a string (newline-separated) or array
  const lines = Array.isArray(mappingDetail)
    ? mappingDetail
    : String(mappingDetail || "").split(/\n\n+/).filter(Boolean);

  return [
    subHeading("Interpretation"),
    new Paragraph({ children: [run(interpretation, { size: 19, color: C.mid })], spacing: { after: 120 } }),
    subHeading("Mapping Summary"),
    ...lines.map(line => new Paragraph({
      children: [run(String(line).replace(/\*\*/g, ""), { size: 19, color: C.mid })],
      spacing: { after: 80 },
    })),
    subHeading("Differences"),
    new Paragraph({ children: [run(differences, { size: 19, color: C.mid })], spacing: { after: 120 } }),
    subHeading("Overall Opinion"),
    new Paragraph({ children: [run(opinion, { size: 19, color: C.mid })], spacing: { after: 160 } }),
  ];
}

// ── Excerpt item ──────────────────────────────────────────────────────────
function excerptItem(num, ref, heading, bodyLines, W = PG.W) {
  const labelW = 1100;
  const refW   = W - labelW;
  const bodyChildren = [];
  if (heading) {
    bodyChildren.push(new Paragraph({
      children: [new TextRun({ text: heading, font: "Arial", size: 16,
        bold: true, color: C.navy, characterSpacing: 30 })],
      spacing: { after: 80 },
    }));
  }
  bodyLines.forEach(line => {
    bodyChildren.push(new Paragraph({
      children: [new TextRun({ text: String(line), font: "Courier New", size: 15, color: C.mid })],
      spacing: { after: 60 },
    }));
  });
  return new Table({
    width: { size: W, type: WidthType.DXA },
    columnWidths: [labelW, refW],
    borders: noBorders,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            borders: { top: solidBorder(C.rule,4), bottom: solidBorder(C.rule,4),
              left: solidBorder(C.rule,4), right: noBorder },
            shading: shade(C.surfaceAlt), margins: CM, width: { size: labelW, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: "Excerpt " + num, font: "Arial",
                size: 17, bold: true, color: C.navy, characterSpacing: 30 })],
              spacing: { after: 0 },
            })],
          }),
          new TableCell({
            borders: { top: solidBorder(C.rule,4), bottom: solidBorder(C.rule,4),
              left: noBorder, right: solidBorder(C.rule,4) },
            shading: shade(C.surfaceAlt), margins: CM, width: { size: refW, type: WidthType.DXA },
            children: [new Paragraph({
              children: [new TextRun({ text: ref, font: "Courier New", size: 15, color: C.muted })],
              alignment: AlignmentType.RIGHT, spacing: { after: 0 },
            })],
          }),
        ],
      }),
      new TableRow({
        children: [new TableCell({
          columnSpan: 2,
          borders: { top: noBorder, bottom: solidBorder(C.rule,4),
            left: solidBorder(C.rule,4), right: solidBorder(C.rule,4) },
          shading: shade("FAFAFA"), margins: CM, width: { size: W, type: WidthType.DXA },
          children: bodyChildren,
        })],
      }),
    ],
  });
}

// ── Feature block (landscape claim chart) ────────────────────────────────
function featureBlock(num, featureText, disclosure, essentiality,
                      analysisChildren, excerptTables, W = PGL.W) {
  const verdictText = disclosure + "  ·  " + essentiality;
  const colW = Math.floor(W / 2);

  const leftChildren = [
    new Paragraph({
      children: [new TextRun({ text: "ANALYSIS", font: "Arial", size: 17,
        bold: true, color: C.muted, characterSpacing: 60 })],
      border: { bottom: solidBorder(C.rule, 4) },
      spacing: { after: 140, before: 0 },
    }),
    ...analysisChildren,
  ];

  const rightChildren = [
    new Paragraph({
      children: [new TextRun({ text: "CITED STANDARD EXCERPTS", font: "Arial",
        size: 17, bold: true, color: C.muted, characterSpacing: 60 })],
      border: { bottom: solidBorder(C.rule, 4) },
      spacing: { after: 140, before: 0 },
    }),
    ...excerptTables.flatMap(t => [t, emptyPara()]),
  ];

  return [
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [400, W - 400],
      borders: noBorders,
      rows: [new TableRow({
        children: [
          new TableCell({
            borders: noBorders, shading: shade(C.navy),
            margins: { top: 120, bottom: 120, left: 160, right: 80 },
            width: { size: 400, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              children: [new TextRun({ text: String(num), font: "Arial",
                size: 28, bold: true, color: C.white })],
              alignment: AlignmentType.CENTER, spacing: { after: 0 },
            })],
          }),
          new TableCell({
            borders: noBorders, shading: shade(C.navy),
            margins: { top: 120, bottom: 120, left: 80, right: 160 },
            width: { size: W - 400, type: WidthType.DXA }, verticalAlign: VerticalAlign.CENTER,
            children: [new Paragraph({
              children: [new TextRun({ text: featureText, font: "Arial",
                size: 19, color: "DDDDDD", italics: true })],
              spacing: { after: 0 },
            })],
          }),
        ],
      })],
    }),
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [W],
      borders: noBorders,
      rows: [new TableRow({
        children: [new TableCell({
          borders: { top: noBorder, bottom: solidBorder(C.rule, 4), left: noBorder, right: noBorder },
          shading: shade(C.amberBg),
          margins: { top: 80, bottom: 80, left: 160, right: 160 },
          width: { size: W, type: WidthType.DXA },
          children: [new Paragraph({
            children: [new TextRun({ text: verdictText, font: "Arial",
              size: 17, bold: true, color: C.amberText, characterSpacing: 30 })],
            spacing: { after: 0 },
          })],
        })],
      })],
    }),
    new Table({
      width: { size: W, type: WidthType.DXA },
      columnWidths: [colW, W - colW],
      borders: noBorders,
      rows: [new TableRow({
        children: [
          new TableCell({
            borders: { top: noBorder, bottom: solidBorder(C.rule, 4),
              left: solidBorder(C.rule, 4), right: solidBorder(C.rule, 4) },
            shading: shade(C.white), margins: CMW,
            width: { size: colW, type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP,
            children: leftChildren,
          }),
          new TableCell({
            borders: { top: noBorder, bottom: solidBorder(C.rule, 4),
              left: noBorder, right: solidBorder(C.rule, 4) },
            shading: shade(C.white), margins: CMW,
            width: { size: W - colW, type: WidthType.DXA }, verticalAlign: VerticalAlign.TOP,
            children: rightChildren,
          }),
        ],
      })],
    }),
    emptyPara(),
  ];
}

// ── Header / Footer ───────────────────────────────────────────────────────
function makeHeader(contentW) {
  return new Header({
    children: [
      new Table({
        width: { size: contentW, type: WidthType.DXA },
        columnWidths: [2800, contentW - 2800],
        borders: noBorders,
        rows: [new TableRow({
          children: [
            new TableCell({
              borders: noBorders, shading: shade(C.navy),
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              width: { size: 2800, type: WidthType.DXA },
              children: [new Paragraph({
                children: [new ImageRun({
                  type: "png", data: logoData,
                  transformation: { width: 160, height: 52 },
                  altText: { title: "IPMIND", description: "IPMIND logo", name: "logo" },
                })],
                spacing: { after: 0 },
              })],
            }),
            new TableCell({
              borders: noBorders, shading: shade(C.navy),
              margins: { top: 80, bottom: 80, left: 120, right: 120 },
              width: { size: contentW - 2800, type: WidthType.DXA },
              verticalAlign: VerticalAlign.CENTER,
              children: [new Paragraph({
                children: [new TextRun({ text: "CONFIDENTIAL", font: "Arial",
                  size: 16, color: "888888", characterSpacing: 80 })],
                alignment: AlignmentType.RIGHT, spacing: { after: 0 },
              })],
            }),
          ],
        })],
      }),
      new Paragraph({
        children: [run("")],
        border: { bottom: solidBorder(C.orange, 12) },
        spacing: { after: 0, before: 0 },
      }),
    ],
  });
}

function makeFooter() {
  return new Footer({
    children: [new Paragraph({
      children: [
        new TextRun({ text: "ipmind.ai", font: "Arial", size: 16, color: C.muted }),
        new TextRun({ text: "\t", font: "Arial", size: 16 }),
        new TextRun({ text: "Page ", font: "Arial", size: 16, color: C.muted }),
        new TextRun({ children: [PageNumber.CURRENT], font: "Arial", size: 16, color: C.muted }),
        new TextRun({ text: " of ", font: "Arial", size: 16, color: C.muted }),
        new TextRun({ children: [PageNumber.TOTAL_PAGES], font: "Arial", size: 16, color: C.muted }),
      ],
      border: { top: solidBorder(C.rule, 4) },
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      spacing: { before: 120, after: 0 },
    })],
  });
}

// ── Disclaimer ────────────────────────────────────────────────────────────
function disclaimerSection() {
  const items = [
    ["Preliminary and Informational Nature:", "The present work product was generated using a prototype AI model and is provided for informational purposes only. It does not constitute a legal or technical opinion regarding the essentiality or non-essentiality of any patent claim to any technical standard."],
    ["Scope of Analysis:", "The analysis is limited to the individual patent claim(s) identified in the chart and does not take into account the full patent specification, including the description and drawings."],
    ["Referencing of Standards:", "Where citations to section numbers, table numbers, or figure numbers in a technical standard are provided, they are included for convenience only and should not be relied upon as authoritative without verification against the official version of the standard."],
    ["Interpretation of Standards:", "References to technical standards are based on publicly available documents. Figures and diagrams from such standards are not reproduced; instead, any associated visual content is paraphrased using descriptive language."],
    ["Subjectivity of Essentiality:", "Determinations of potential alignment between a patent claim and a standard may depend on how specific terms or functional steps are construed. This assessment is inherently interpretive and does not reflect a consensus view or judicial determination."],
    ["Implementation Considerations:", "The presence of a feature in a standard does not imply that all compliant implementations necessarily use that feature."],
    ["Alternative Solutions:", "Standards may include multiple options or alternative techniques to achieve similar functionality. A given patent claim may correspond to one such option, but not to others that are also compliant with the standard."],
    ["Legal Proceedings:", "In the context of litigation, essentiality determinations typically require expert testimony, claim construction under applicable law, and examination of implementation evidence. The present assessment should not be relied upon for litigation or licensing negotiation without further professional review."],
  ];
  return [
    ...sectionHeading("Disclaimer"),
    new Table({
      width: { size: PG.W, type: WidthType.DXA },
      columnWidths: [PG.W],
      borders: noBorders,
      rows: [new TableRow({
        children: [new TableCell({
          borders: {
            top: solidBorder(C.rule,4), bottom: solidBorder(C.rule,4),
            left: solidBorder(C.rule,4), right: solidBorder(C.rule,4),
          },
          shading: shade(C.surfaceAlt), margins: CMW,
          width: { size: PG.W, type: WidthType.DXA },
          children: items.map((item, i) => new Paragraph({
            children: [
              new TextRun({ text: (i+1) + ".  " + item[0] + "  ",
                font: "Arial", size: 19, bold: true, color: C.mid }),
              new TextRun({ text: item[1], font: "Arial", size: 19, color: C.muted }),
            ],
            spacing: { after: 120 },
          })),
        })],
      })],
    }),
  ];
}

// ── Parse excerpt markdown string ─────────────────────────────────────────
function parseExcerpt(excStr) {
  // Excerpt number — handle "1", "1.", "1  " etc.
  const numMatch = excStr.match(/\*\*Excerpt_Number:\*\*\s*([^\n\s]+)/);
  const num      = numMatch ? numMatch[1].replace(/\.$/, "") : "?";

  // Extract body after "Excerpt_Text:** Excerpt:" — grab everything to end, strip trailing ---
  const textMatch = excStr.match(/\*\*Excerpt_Text:\*\*\s*Excerpt:[ \t]*\n([\s\S]+)/);
  const rawBody   = textMatch
    ? textMatch[1].replace(/\n---[ \t]*$/, "").trim()
    : excStr;

  // Reference: match both **bold** and plain formats, across one or two lines
  // Pattern 1: Reference:\n**text**
  // Pattern 2: Reference:\nplain text
  // Pattern 3: Reference: plain text (inline)
  const refMatch =
    rawBody.match(/Reference:[ \t]*\n\*\*([^*\n]+)\*\*/) ||
    rawBody.match(/Reference:[ \t]*\n([^\n*][^\n]+)/)     ||
    rawBody.match(/Reference:[ \t]+([^\n]+)/);
  const ref = refMatch ? refMatch[1].trim() : "";

  // Strip the reference block from body
  const bodyStripped = rawBody
    .replace(/\nReference:[ \t]*\n\*\*[^*]+\*\*[ \t]*/g, "")
    .replace(/\nReference:[ \t]*\n[^\n]+[ \t]*/g, "")
    .replace(/\nReference:[ \t]+[^\n]+/g, "")
    .trim();

  // Section heading from ## line
  const h2Match = bodyStripped.match(/^##[ \t]+(.+)/m);
  const heading = h2Match ? h2Match[1].trim() : "";

  // Split into lines, skip top-level # headings and blank lines
  const bodyLines = bodyStripped
    .split("\n")
    .filter(l => !l.trim().startsWith("# ") && l.trim() !== "")
    .map(l => l.trim());

  return { num, ref, heading, bodyLines };
}

// ── Limitations: first line = label, rest = body ──────────────────────────
function parseLimitations(str) {
  const lines = (str || "").split("\n");
  const label = lines[0].trim();
  const body  = lines.slice(1).join("\n").replace(/^\s*\n/, "").trim();
  return { label, body };
}

// ═════════════════════════════════════════════════════════════════════════
// DOCUMENT BUILDER
// ═════════════════════════════════════════════════════════════════════════

// ── Restricted Use Notice — Word page ─────────────────────────────────────
const RESTRICTED_NOTICE_TEXT =
  "This report is confidential and provided solely for internal use in connection " +
  "with patent licensing, portfolio evaluation, or standards-related strategy. It must " +
  "not be published, posted, or circulated to any third party without IP Mind\u2019s prior " +
  "written consent. Where disclosure to a counterparty is necessary, the report may be " +
  "shared in full or in part provided the counterparty is bound by a written " +
  "confidentiality undertaking that places equivalent restrictions on use and further " +
  "distribution, and that requires attribution of IP Mind\u2019s authorship to be retained. " +
  "The recipient must not use this report to replicate, benchmark, or train models " +
  "intended to reproduce IP Mind\u2019s methodology or outputs, or to develop competing " +
  "analysis products or services.";

function restrictedNoticePage() {
  return [
    new Paragraph({
      children: [new PageBreak()],
      spacing: { after: 0 },
    }),
    new Paragraph({
      children: [new TextRun({ text: "RESTRICTED USE NOTICE", font: "Arial",
        size: 17, bold: true, color: C.orange, characterSpacing: 80 })],
      spacing: { before: 480, after: 240 },
      border: { bottom: solidBorder(C.orange, 8) },
    }),
    new Table({
      width: { size: PG.W, type: WidthType.DXA },
      columnWidths: [60, PG.W - 60],
      borders: noBorders,
      rows: [new TableRow({
        children: [
          new TableCell({
            borders: { top: noBorder, bottom: noBorder,
              left: solidBorder(C.orange, 12), right: noBorder },
            shading: shade(C.amberBg),
            margins: { top: 0, bottom: 0, left: 0, right: 0 },
            width: { size: 60, type: WidthType.DXA },
            children: [emptyPara()],
          }),
          new TableCell({
            borders: noBorders,
            shading: shade(C.amberBg),
            margins: { top: 200, bottom: 200, left: 240, right: 240 },
            width: { size: PG.W - 60, type: WidthType.DXA },
            children: [
              new Paragraph({
                children: [new TextRun({ text: RESTRICTED_NOTICE_TEXT,
                  font: "Arial", size: 19, color: C.amberText })],
                spacing: { after: 0 },
              }),
            ],
          }),
        ],
      })],
    }),
    new Paragraph({
      children: [new PageBreak()],
      spacing: { after: 0 },
    }),
  ];
}

async function buildDocument(data, meta, restricted) {
  const patentNumber  = data.Patent_Number || meta.Patent_Number || "Unknown";
  const title         = data.Title         || meta.Title         || "Patent Analysis Report";
  const owner         = data.Owner         || meta.Owner         || "";
  const standard      = data.Standard      || meta.Standard      || "";
  const claimNumber   = data.Claim_Number  || "";
  const claimText     = data.Claim         || "";
  const claimCategory = data.Claim_Category|| "";
  const pctMapped     = data.Mapped_Percentage || "";
  const pctWeighted   = data["Mapped_Percentage_(Weighted)"] || "";
  const essDecision   = data.Essentiality_Conclusion || "";
  const opinion       = data.Summary       || "";
  const mappingItems  = data.Mapping_Summary || [];
  const charts        = data.Claim_Charts  || [];
  const { label: limLabel, body: limBody } = parseLimitations(data["Limitation(s)"] || "");
  const claimLabel    = claimNumber + " \u2014 " + claimCategory + " Claim";

  // ── Section 1: Portrait — Identity, Summary, Mapping ──────────────────
  const section1Children = [
    // Title area
    new Paragraph({
      children: [
        new TextRun({ text: patentNumber + "  ", font: "Arial", size: 17, bold: true,
          color: C.orange, characterSpacing: 40 }),
        new TextRun({ text: standard + "  ", font: "Arial", size: 17, bold: true,
          color: C.navy, characterSpacing: 40 }),
        new TextRun({ text: claimNumber + " \u00B7 " + claimCategory,
          font: "Arial", size: 17, bold: true, color: C.navy, characterSpacing: 40 }),
        ...(restricted ? [new TextRun({ text: "   \u2014   CONFIDENTIAL \u00B7 RESTRICTED USE  \u2014  See page 2",
          font: "Arial", size: 15, color: C.amberText, characterSpacing: 20 })] : []),
      ],
      spacing: { before: 320, after: 120 },
    }),
    new Paragraph({
      children: [new TextRun({ text: title, font: "Georgia", size: 52, bold: true, color: C.navy })],
      spacing: { after: 280 },
    }),
    // Identity grid
    new Table({
      width: { size: PG.W, type: WidthType.DXA },
      columnWidths: [Math.floor(PG.W/3), Math.floor(PG.W/3), PG.W - Math.floor(PG.W/3)*2],
      borders: noBorders,
      rows: [new TableRow({
        children: [
          { label: "Patent Number", value: patentNumber },
          { label: "Owner",         value: owner },
          { label: "Standard",      value: standard },
        ].map(cell => new TableCell({
          borders: { top: solidBorder(C.rule,4), bottom: solidBorder(C.rule,4),
            left: noBorder, right: solidBorder(C.rule,4) },
          shading: shade(C.white), margins: CMW,
          width: { size: Math.floor(PG.W/3), type: WidthType.DXA },
          children: [
            new Paragraph({ children: [new TextRun({ text: cell.label.toUpperCase(),
              font: "Arial", size: 15, bold: true, color: C.muted, characterSpacing: 40 })],
              spacing: { after: 40 } }),
            new Paragraph({ children: [new TextRun({ text: cell.value,
              font: "Arial", size: 20, bold: true, color: C.navy })],
              spacing: { after: 0 } }),
          ],
        }))
      })],
    }),
    emptyPara(),
    claimBlock(claimLabel, claimText),
    emptyPara(),
    ...(restricted ? restrictedNoticePage() : []),
    ...sectionHeading("Executive Summary"),
    summaryCardTable([
      { label: "Claim Number",         value: claimNumber },
      { label: "Claim Category",       value: claimCategory },
      { label: "Pct. Mapped",          value: pctMapped },
      { label: "Essentiality Decision",value: essDecision, highlight: true, small: true },
    ]),
    emptyPara(),
    summaryCardTable([
      { label: "Weighted Mapping", value: pctWeighted },
      { label: "Limitations",      value: limLabel, small: true },
    ]),
    emptyPara(),
    // Opinion box
    new Table({
      width: { size: PG.W, type: WidthType.DXA },
      columnWidths: [PG.W],
      borders: noBorders,
      rows: [new TableRow({
        children: [new TableCell({
          borders: { top: solidBorder(C.rule,4), bottom: solidBorder(C.rule,4),
            left: solidBorder(C.rule,4), right: solidBorder(C.rule,4) },
          shading: shade(C.white), margins: CMW,
          width: { size: PG.W, type: WidthType.DXA },
          children: [
            new Paragraph({ children: [new TextRun({ text: "Opinion",
              font: "Georgia", size: 28, bold: true, color: C.navy })],
              spacing: { after: 120 } }),
            new Paragraph({ children: [run(opinion, { size: 19, color: C.mid })],
              spacing: { after: 160 } }),
            new Paragraph({ children: [new TextRun({ text: "Limitations Detail",
              font: "Arial", size: 15, bold: true, color: C.muted, characterSpacing: 40 })],
              shading: shade(C.surfaceAlt), spacing: { after: 80, before: 120 } }),
            new Paragraph({ children: [run(limBody, { size: 19, color: C.mid })],
              spacing: { after: 0 } }),
          ],
        })],
      })],
    }),
    emptyPara(),
    ...sectionHeading("Mapping Summary"),
    ...mappingItems.flatMap((item, i) => [
      mappingItem(i + 1, item.Key_Feature, item.Conclusions, item.Brief_Rationale),
      emptyPara(),
    ]),
  ];

  // ── Section 2: Landscape — Claim Chart ────────────────────────────────
  const section2Children = [
    new Paragraph({
      children: [new TextRun({ text: "Claim Chart", font: "Georgia",
        size: 32, bold: true, color: C.navy })],
      spacing: { before: 480, after: 160 },
      border: { bottom: solidBorder(C.rule, 4) },
    }),
    ...charts.flatMap(chart => {
      const feat = chart.Claim_Feature || {};
      const dec  = chart.Decision      || {};
      const ana  = chart.Analysis      || {};
      const excRaw = chart.Cited_Excerpts || [];

      const colW   = Math.floor(PGL.W / 2);
      const innerW = colW - (CMW.left + CMW.right);

      const analysisChildren = [
        ...analysisParagraphs(
          ana.Interpretation  || "",
          ana.Mapping_Summary || "",
          ana.Differences     || "",
          ana.Overall_Opinion || ""
        ),
        justificationPanel(dec.Justification || "", innerW),
      ];

      const excerptTables = excRaw.map((excStr, i) => {
        const exc = parseExcerpt(excStr);
        return excerptItem(exc.num, exc.ref, exc.heading, exc.bodyLines, innerW);
      });

      return featureBlock(
        feat.Index || (charts.indexOf(chart) + 1),
        feat.Text  || "",
        dec.Disclosure || "",
        dec.Essentiality_Classification || "",
        analysisChildren,
        excerptTables,
        PGL.W
      );
    }),
  ];

  // ── Section 3: Portrait — Disclaimer ──────────────────────────────────
  const section3Children = [...disclaimerSection(), emptyPara()];

  const doc = new Document({
    sections: [
      {
        properties: {
          type: SectionType.NEXT_PAGE,
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        headers: { default: makeHeader(PG.W) },
        footers: { default: makeFooter() },
        children: section1Children,
      },
      {
        properties: {
          type: SectionType.NEXT_PAGE,
          page: {
            size: { width: 11906, height: 16838, orientation: PageOrientation.LANDSCAPE },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        headers: { default: makeHeader(PGL.W) },
        footers: { default: makeFooter() },
        children: section2Children,
      },
      {
        properties: {
          type: SectionType.NEXT_PAGE,
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
          },
        },
        headers: { default: makeHeader(PG.W) },
        footers: { default: makeFooter() },
        children: section3Children,
      },
    ],
  });

  return Packer.toBuffer(doc);
}

// ═════════════════════════════════════════════════════════════════════════
// EXPRESS ROUTES
// ═════════════════════════════════════════════════════════════════════════

// Health check
app.get("/", (req, res) => {
  res.json({ status: "ok", service: "ipmind-docx-service" });
});

// Main endpoint — POST the IPMIND analysis JSON
// Optional query params: patent, title, owner, standard (override META)
app.post("/generate", async (req, res) => {
  try {
    let body = req.body;

    // Accept both array [ {...} ] and plain object { ... }
    const data = Array.isArray(body) ? body[0] : body;

    // Patent metadata can also be passed as query params
    const meta = {
      Patent_Number: req.query.patent  || "",
      Title:         req.query.title   || "",
      Owner:         req.query.owner   || "",
      Standard:      req.query.standard|| "",
    };

    const restricted = req.query.restricted === "true" || data.Restricted_Use === true;
    const buf      = await buildDocument(data, meta, restricted);
    const safeName = (data.Patent_Number || meta.Patent_Number || "report")
      .replace(/[^A-Za-z0-9_-]/g, "_");
    const filename = safeName + "_report.docx";

    res.setHeader("Content-Type",
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    res.setHeader("Content-Disposition", 'attachment; filename="' + filename + '"');
    res.setHeader("Content-Length", buf.length);
    res.send(buf);

  } catch (err) {
    console.error("Error generating docx:", err);
    res.status(500).json({ error: err.message });
  }
});

// ═════════════════════════════════════════════════════════════════════════
// HTML GENERATOR
// ═════════════════════════════════════════════════════════════════════════

const RESTRICTED_NOTICE_HTML =
  "This report is confidential and provided solely for internal use in connection " +
  "with patent licensing, portfolio evaluation, or standards-related strategy. It must " +
  "not be published, posted, or circulated to any third party without IP Mind\u2019s prior " +
  "written consent. Where disclosure to a counterparty is necessary, the report may be " +
  "shared in full or in part provided the counterparty is bound by a written " +
  "confidentiality undertaking that places equivalent restrictions on use and further " +
  "distribution, and that requires attribution of IP Mind\u2019s authorship to be retained. " +
  "The recipient must not use this report to replicate, benchmark, or train models " +
  "intended to reproduce IP Mind\u2019s methodology or outputs, or to develop competing " +
  "analysis products or services.";

function buildHtml(data, meta, restricted) {
  const patentNumber  = data.Patent_Number  || meta.Patent_Number  || "";
  const title         = data.Title          || meta.Title          || "";
  const owner         = data.Owner          || meta.Owner          || "";
  const standard      = data.Standard       || meta.Standard       || "";
  const claimNumber   = data.Claim_Number   || "";
  const claimText     = data.Claim          || "";
  const claimCategory = data.Claim_Category || "";
  const pctMapped     = data.Mapped_Percentage || "";
  const pctWeighted   = data["Mapped_Percentage_(Weighted)"] || "";
  const essDecision   = data.Essentiality_Conclusion || "";
  const opinion       = data.Summary        || "";
  const mappingItems  = data.Mapping_Summary || [];
  const charts        = data.Claim_Charts   || [];

  function parseLimitationsH(str) {
    const lines = (str || "").split("\n");
    const label = lines[0].trim();
    const body  = lines.slice(1).join("\n").replace(/^\s*\n/, "").trim();
    return { label, body };
  }
  const { label: limLabel, body: limBody } = parseLimitationsH(data["Limitation(s)"] || "");
  const claimLabel = `${claimNumber} \u2014 ${claimCategory} Claim`;

  function esc(str) {
    return String(str || "")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  function renderInline(str) {
    return String(str || "")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>")
      .replace(/(?<!\*)\*(?!\*)(.+?)(?<!\*)\*(?!\*)/g, "<em>$1</em>");
  }

  function renderMdTable(tableLines) {
    let thead = "", tbody = "";
    tableLines.forEach((line, idx) => {
      if (/^\|[\s\-:|]+\|$/.test(line.trim())) return;
      const cells = line.split("|").slice(1, -1);
      if (idx === 0) {
        thead = "<thead><tr>" + cells.map(c => `<th>${renderInline(c.trim())}</th>`).join("") + "</tr></thead>";
      } else {
        tbody += "<tr>" + cells.map(c => `<td>${renderInline(c.trim())}</td>`).join("") + "</tr>";
      }
    });
    return `<table class="exc-table">${thead}<tbody>${tbody}</tbody></table>`;
  }

  function renderMdBlock(md) {
    if (!md) return "";
    const lines = md.split("\n");
    let out = "", i = 0;
    while (i < lines.length) {
      const line = lines[i];
      const trimmed = line.trim();
      if (trimmed.startsWith("|")) {
        const tableLines = [];
        while (i < lines.length && lines[i].trim().startsWith("|")) { tableLines.push(lines[i]); i++; }
        out += renderMdTable(tableLines); continue;
      }
      if (trimmed.startsWith("## ")) { out += `<p class="exc-subhead">${renderInline(trimmed.slice(3))}</p>`; i++; continue; }
      if (trimmed.startsWith("# "))  { i++; continue; }
      if (/^[\t\s]*[\u2014\u2013-]/.test(line) && trimmed.startsWith("\u2014")) {
        out += '<ul class="exc-list">';
        while (i < lines.length && lines[i].trim().startsWith("\u2014")) {
          const text = lines[i].trim().replace(/^[\u2014]\s*/, "");
          out += `<li>${renderInline(text)}</li>`; i++;
        }
        out += "</ul>"; continue;
      }
      if (/^ {2,}/.test(line) && trimmed !== "") { out += `<p class="exc-indent">${renderInline(trimmed)}</p>`; i++; continue; }
      if (trimmed === "") { i++; continue; }
      out += `<p>${renderInline(trimmed)}</p>`; i++;
    }
    return out;
  }

  function renderAnalysisBlock(str) {
    if (!str) return "<p></p>";
    return str.split(/\n\n+/).map(para => {
      para = para.trim();
      if (!para) return "";
      const subLines = para.split("\n").filter(l => l.trim());
      if (subLines.length > 1) return subLines.map(l => `<p>${renderInline(l.trim())}</p>`).join("");
      return `<p>${renderInline(para)}</p>`;
    }).join("");
  }

  function essClasses(decision) {
    const d = (decision || "").toLowerCase();
    if (d.includes("not essential"))  return { card: "", value: "red", dot: "dot-red", verdict: "red", badge: "badge-red" };
    if (d.includes("conditional"))    return { card: "highlight", value: "amber", dot: "dot-amber", verdict: "amber", badge: "badge-amber" };
    if (d.includes("essential"))      return { card: "highlight-green", value: "green", dot: "dot-green", verdict: "green", badge: "badge-green" };
    return { card: "highlight", value: "amber", dot: "dot-amber", verdict: "amber", badge: "badge-amber" };
  }

  // Reuse the improved parseExcerpt but return bodyHtml instead of bodyLines
  function parseExcerptHtml(excStr) {
    const numMatch  = excStr.match(/\*\*Excerpt_Number:\*\*\s*([^\n\s]+)/);
    const num       = numMatch ? numMatch[1] : "?";
    const textMatch = excStr.match(/\*\*Excerpt_Text:\*\*\s*Excerpt:[ \t]*\n([\s\S]+)/);
    const rawBody   = textMatch ? textMatch[1].replace(/\n---[ \t]*$/, "").trim() : excStr;
    const refMatch  =
      rawBody.match(/Reference:[ \t]*\n\*\*([^*\n]+)\*\*/) ||
      rawBody.match(/Reference:[ \t]*\n([^\n*][^\n]+)/)     ||
      rawBody.match(/Reference:[ \t]+([^\n]+)/);
    const ref = refMatch ? refMatch[1].trim() : "";
    const bodyStripped = rawBody
      .replace(/\nReference:[ \t]*\n\*\*[^*]+\*\*[ \t]*/g, "")
      .replace(/\nReference:[ \t]*\n[^\n]+[ \t]*/g, "")
      .replace(/\nReference:[ \t]+[^\n]+/g, "")
      .trim();
    const h2Match = bodyStripped.match(/^##[ \t]+(.+)/m);
    const heading = h2Match ? h2Match[1].trim() : "";
    const bodyHtml = renderMdBlock(bodyStripped);
    return { num, ref, heading, bodyHtml };
  }

  const ec = essClasses(essDecision);

  function buildMappingSummaryHtml(items) {
    return items.map(item => {
      const bc = essClasses(item.Conclusions || "").badge;
      const conclusions = item.Conclusions || "";
      const badgeLabel = conclusions.includes("|") ? conclusions.split("|")[0].trim() : conclusions;
      return `
        <div class="mapping-item">
          <div class="mapping-item-header">
            <div class="feat-num">${esc(item.Index)}</div>
            <div class="feat-text">${esc(item.Key_Feature)}</div>
            <div><span class="badge ${bc}">${esc(badgeLabel)}</span></div>
          </div>
          <div class="mapping-item-body">
            <p><strong>Conclusion:</strong> ${esc(conclusions)}</p>
            <p><strong>Brief Rationale:</strong> ${esc(item.Brief_Rationale)}</p>
          </div>
        </div>`;
    }).join("\n");
  }

  function buildClaimChartHtml(charts) {
    return charts.map(chart => {
      const feat  = chart.Claim_Feature || {};
      const dec   = chart.Decision      || {};
      const ana   = chart.Analysis      || {};
      const excRaw = chart.Cited_Excerpts || [];
      const disclosure = dec.Disclosure || "";
      const essClass   = dec.Essentiality_Classification || "";
      const justification = dec.Justification || "";
      const fc = essClasses(essClass);
      const parsedExcs = excRaw.map(parseExcerptHtml);
      const excItemsHtml = parsedExcs.map(exc => `
              <div class="excerpt-item">
                <div class="excerpt-item-header">
                  <span class="exc-num">Excerpt ${esc(exc.num)}</span>
                  <span class="exc-ref">${esc(exc.ref)}</span>
                </div>
                <div class="excerpt-item-body">
                  ${exc.heading ? `<h4>${esc(exc.heading)}</h4>` : ""}
                  ${exc.bodyHtml}
                </div>
              </div>`).join("\n");
      return `
      <div class="claim-feature-block">
        <div class="cfb-header">
          <div class="feat-num">${esc(feat.Index)}</div>
          <div class="feat-title">${esc(feat.Text)}</div>
        </div>
        <div class="cfb-verdict">
          <div class="verdict-item"><div class="verdict-dot ${fc.dot}"></div><span class="${fc.verdict}">${esc(disclosure)}</span></div>
          <span class="verdict-sep">&middot;</span>
          <div class="verdict-item"><div class="verdict-dot ${fc.dot}"></div><span class="${fc.verdict}">${esc(essClass)}</span></div>
        </div>
        <div class="cfb-body">
          <div class="cfb-col">
            <h4>Analysis</h4>
            <div class="sub-heading">Interpretation</div>${renderAnalysisBlock(ana.Interpretation)}
            <div class="sub-heading">Mapping Summary</div>${renderAnalysisBlock(ana.Mapping_Summary)}
            <div class="sub-heading">Differences</div>${renderAnalysisBlock(ana.Differences)}
            <div class="sub-heading">Overall Opinion</div>${renderAnalysisBlock(ana.Overall_Opinion)}
            <div class="justification-panel">
              <div class="j-label">Essentiality Justification</div>
              <p>${esc(justification)}</p>
            </div>
          </div>
          <div class="cfb-col">
            <h4>Cited Standard Excerpts</h4>
            <div class="excerpts-section">
              <button class="excerpt-toggle" onclick="toggleExcerpts(this)" aria-expanded="false">
                <span class="toggle-left"><span>Standard Excerpts</span><span class="excerpt-count">${parsedExcs.length}</span></span>
                <svg class="chevron" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="6 9 12 15 18 9"/></svg>
              </button>
              <div class="excerpt-body">${excItemsHtml}</div>
            </div>
          </div>
        </div>
      </div>`;
    }).join("\n");
  }

  const LOGO_SVG = `<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" version="1.1" height="36" viewBox="0 0 3144.8497854077254 1027.5281652360513"><g transform="scale(7.242489270386266) translate(10, 10)"><defs id="SvgjsDefs1027"/><g id="SvgjsG1028" featureKey="symbolGroupContainer" transform="matrix(1.16515289568328,0,0,1.16515289568328,0.000007264315552150539,0.000007264315552150539)" fill="#fff"><path d="M52.3 104.6a52.3 52.3 0 1 1 52.3-52.3 52.4 52.4 0 0 1-52.3 52.3zm0-102.3a50 50 0 1 0 50 50 50 50 0 0 0-50-50z"/></g><g id="SvgjsG1029" featureKey="2ou6gm-0" transform="matrix(0.9971509971509972,0,0,0.9971509971509972,264.8062678062678,-335.4786324786325)" fill="#fff"><path d="M-167.5,390.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-165.5,389.6-166.4,390.5-167.5,390.5z M-177.5,428.5c-2.2,0-4-1.8-4-4s1.8-4,4-4c2.2,0,4,1.8,4,4S-175.3,428.5-177.5,428.5z M-177.5,410.5c-2.2,0-4-1.8-4-4s1.8-4,4-4c2.2,0,4,1.8,4,4S-175.3,410.5-177.5,410.5z M-177.5,392.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-173.5,390.7-175.3,392.5-177.5,392.5z M-177.5,374.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-173.5,372.7-175.3,374.5-177.5,374.5z M-194.5,414.5c-3.9,0-7-3.1-7-7c0-3.9,3.1-7,7-7c3.9,0,7,3.1,7,7C-187.5,411.4-190.6,414.5-194.5,414.5z M-194.5,394.5c-3.9,0-7-3.1-7-7c0-3.9,3.1-7,7-7c3.9,0,7,3.1,7,7C-187.5,391.4-190.6,394.5-194.5,394.5z M-195.5,374.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-191.5,372.7-193.3,374.5-195.5,374.5z M-195.5,362.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-193.5,361.6-194.4,362.5-195.5,362.5z M-214.5,414.5c-3.9,0-7-3.1-7-7c0-3.9,3.1-7,7-7s7,3.1,7,7C-207.5,411.4-210.6,414.5-214.5,414.5z M-214.5,394.5c-3.9,0-7-3.1-7-7c0-3.9,3.1-7,7-7s7,3.1,7,7C-207.5,391.4-210.6,394.5-214.5,394.5z M-213.5,374.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-209.5,372.7-211.3,374.5-213.5,374.5z M-213.5,362.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-211.5,361.6-212.4,362.5-213.5,362.5z M-231.5,374.5c-2.2,0-4-1.8-4-4c0-2.2,1.8-4,4-4c2.2,0,4,1.8,4,4C-227.5,372.7-229.3,374.5-231.5,374.5z M-231.5,384.5c2.2,0,4,1.8,4,4c0,2.2-1.8,4-4,4c-2.2,0-4-1.8-4-4C-235.5,386.3-233.7,384.5-231.5,384.5z M-241.5,408.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-239.5,407.6-240.4,408.5-241.5,408.5z M-241.5,390.5c-1.1,0-2-0.9-2-2c0-1.1,0.9-2,2-2c1.1,0,2,0.9,2,2C-239.5,389.6-240.4,390.5-241.5,390.5z M-231.5,402.5c2.2,0,4,1.8,4,4s-1.8,4-4,4c-2.2,0-4-1.8-4-4S-233.7,402.5-231.5,402.5z M-231.5,420.5c2.2,0,4,1.8,4,4s-1.8,4-4,4c-2.2,0-4-1.8-4-4S-233.7,420.5-231.5,420.5z M-213.5,420.5c2.2,0,4,1.8,4,4c0,2.2-1.8,4-4,4c-2.2,0-4-1.8-4-4C-217.5,422.3-215.7,420.5-213.5,420.5z M-213.5,432.5c1.1,0,2,0.9,2,2c0,1.1-0.9,2-2,2c-1.1,0-2-0.9-2-2C-215.5,433.4-214.6,432.5-213.5,432.5z M-195.5,420.5c2.2,0,4,1.8,4,4c0,2.2-1.8,4-4,4c-2.2,0-4-1.8-4-4C-199.5,422.3-197.7,420.5-195.5,420.5z M-195.5,432.5c1.1,0,2,0.9,2,2c0,1.1-0.9,2-2,2c-1.1,0-2-0.9-2-2C-197.5,433.4-196.6,432.5-195.5,432.5z M-167.5,404.5c1.1,0,2,0.9,2,2c0,1.1-0.9,2-2,2c-1.1,0-2-0.9-2-2C-169.5,405.4-168.6,404.5-167.5,404.5z" style="fill-rule:evenodd;clip-rule:evenodd;"/></g><g id="SvgjsG1030" featureKey="kZnDdN-0" transform="matrix(3.8775259911441498,0,0,3.8775259911441498,137.1154802767278,2.6123700442792526)" fill="#fff"><path d="M2.8906 8.457 c-0.88867 0 -1.6309 -0.72266 -1.6309 -1.6211 c0 -0.88867 0.74219 -1.6113 1.6309 -1.6113 c0.86914 0 1.6113 0.72266 1.6113 1.6113 c0 0.89844 -0.74219 1.6211 -1.6113 1.6211 z M1.4551 20 l0 -10.039 l2.832 0 l0 10.039 l-2.832 0 z M13.0859875 9.766 c2.6465 0 4.834 1.9434 4.834 5.2344 s-2.1875 5.2344 -4.834 5.2344 c-1.3086 0 -2.4805 -0.50781 -3.0762 -1.4258 l0 6.0742 l-2.8125 0 l0 -14.922 l2.666 0 l0.078125 1.3477 c0.55664 -0.99609 1.7773 -1.543 3.1445 -1.543 z M12.4511875 17.9004 c1.4746 0 2.6563 -1.0742 2.6563 -2.9004 s-1.1816 -2.9004 -2.6563 -2.9004 c-1.5039 0 -2.6758 1.1426 -2.6758 2.9004 s1.1719 2.9004 2.6758 2.9004 z M37.129296875 9.766 c2.1484 0 3.5352 1.0938 3.5352 3.1543 l0 7.0801 l-2.8125 0 l0 -6.2793 c0 -1.1816 -0.74219 -1.6992 -1.582 -1.6992 c-1.0059 0 -1.8945 0.57617 -1.8945 2.3145 l0 5.6641 l-2.8418 0 l0 -6.25 c0 -1.2012 -0.72266 -1.7285 -1.6113 -1.7285 c-0.97656 0 -1.8848 0.57617 -1.8848 2.4609 l0 5.5176 l-2.8027 0 l0 -10.039 l2.8027 0 l0 1.1816 c0.66406 -0.83008 1.7871 -1.3086 3.1152 -1.3086 z M44.833959375 8.457 c-0.88867 0 -1.6309 -0.72266 -1.6309 -1.6211 c0 -0.88867 0.74219 -1.6113 1.6309 -1.6113 c0.86914 0 1.6113 0.72266 1.6113 1.6113 c0 0.89844 -0.74219 1.6211 -1.6113 1.6211 z M43.398459375 20 l0 -10.039 l2.832 0 l0 10.039 l-2.832 0 z M55.068346875 9.766 c2.4121 0 3.7402 1.25 3.7402 3.4766 l0 6.7578 l-2.8223 0 l0 -6.1523 c0 -1.3379 -0.83008 -1.8262 -1.7969 -1.8262 c-1.1621 0 -2.2168 0.58594 -2.2363 2.4414 l0 5.5371 l-2.8125 0 l0 -10.039 l2.8125 0 l0 1.1133 c0.70313 -0.83008 1.7871 -1.3086 3.1152 -1.3086 z M68.652325 5 l2.8125 0 l0 15 l-2.666 0 l-0.068359 -1.3086 c-0.57617 0.98633 -1.7871 1.543 -3.1543 1.543 c-2.6465 0 -4.834 -1.9531 -4.834 -5.2344 s2.1973 -5.2344 4.834 -5.2344 c1.3184 0 2.4805 0.49805 3.0762 1.4063 l0 -6.1719 z M66.220725 17.9004 c1.4941 0 2.6563 -1.1426 2.6563 -2.9004 s-1.1719 -2.9102 -2.6563 -2.9102 c-1.4941 0 -2.666 1.1035 -2.666 2.9102 c0 1.7969 1.1719 2.9004 2.666 2.9004 z"/></g></g></svg>`;

  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8" />
  <title>${esc(patentNumber)} \u2013 Patent Analysis Report</title>
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <link rel="preconnect" href="https://fonts.googleapis.com" />
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin />
  <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;600;700&family=Source+Sans+3:wght@300;400;500;600&family=Source+Code+Pro:wght@400;500&display=swap" rel="stylesheet" />
  <style>
    :root{--brand:#ff6734;--brand-light:#fff0eb;--navy:#0f1f38;--ink:#1c1c2e;--mid:#4a4a6a;--muted:#7a7a96;--rule:#e2e2ed;--bg:#fafaf8;--surface:#ffffff;--surface-alt:#f4f4f0;--green:#1a6b4a;--green-bg:#eaf5ef;--amber:#8a5a00;--amber-bg:#fdf5e0;--red:#8a0000;--red-bg:#fdf0f0;--radius:6px;--radius-lg:12px;}
    *,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
    html{scroll-behavior:smooth;}
    body{font-family:'Source Sans 3',sans-serif;font-size:15px;line-height:1.7;color:var(--ink);background:var(--bg);}
    @page{margin:2cm;}
    @media print{.excerpt-toggle{display:none;}.excerpt-body{display:block!important;}}
    .page-wrap{max-width:1040px;margin:0 auto;padding:0 32px 80px;}
    .brand-rule{height:4px;background:var(--brand);}
    .header-bar{background:var(--navy);}
    .header-bar-inner{max-width:1040px;margin:0 auto;padding:28px 32px;display:flex;justify-content:space-between;align-items:center;}
    .header-bar .logo svg{height:32px;width:auto;}
    .header-bar .confidential{font-size:11px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:rgba(255,255,255,.5);border:1px solid rgba(255,255,255,.2);padding:4px 12px;border-radius:2px;}
    .identity{padding:48px 0 40px;border-bottom:1px solid var(--rule);}
    .identity-meta{display:flex;gap:8px;align-items:center;margin-bottom:16px;}
    .pill{display:inline-block;font-size:11px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;padding:3px 10px;border-radius:2px;background:var(--brand-light);color:var(--brand);}
    .pill-navy{background:rgba(15,31,56,.08);color:var(--navy);}
    .pill-restricted{background:var(--amber-bg);color:var(--amber);}
    .restricted-notice{margin:24px 0 0;border-left:3px solid var(--amber);background:var(--amber-bg);padding:18px 24px;border-radius:0 var(--radius) var(--radius) 0;}
    .restricted-label{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--amber);margin-bottom:10px;}
    .restricted-notice p{font-size:13px;line-height:1.75;color:var(--amberText,#8a5a00);}
    .identity h1{font-family:'Playfair Display',serif;font-size:32px;font-weight:600;color:var(--navy);line-height:1.2;margin-bottom:8px;}
    .identity-grid{display:grid;grid-template-columns:repeat(3,1fr);border:1px solid var(--rule);border-radius:var(--radius);overflow:hidden;margin-top:32px;}
    .identity-cell{padding:16px 20px;border-right:1px solid var(--rule);}
    .identity-cell:last-child{border-right:none;}
    .identity-cell .label{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:4px;}
    .identity-cell .value{font-size:14px;font-weight:600;color:var(--navy);}
    .claim-block{margin:36px 0 0;border-left:3px solid var(--brand);background:var(--surface);padding:20px 24px;border-radius:0 var(--radius) var(--radius) 0;}
    .claim-block .claim-label{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--brand);margin-bottom:10px;}
    .claim-block p{font-size:14px;line-height:1.75;color:var(--ink);font-style:italic;}
    .section{margin-top:56px;}
    .section-heading{display:flex;align-items:center;gap:16px;margin-bottom:28px;}
    .section-heading h2{font-family:'Playfair Display',serif;font-size:22px;font-weight:600;color:var(--navy);white-space:nowrap;}
    .section-heading::after{content:'';flex:1;height:1px;background:var(--rule);}
    .summary-cards{display:grid;grid-template-columns:repeat(4,1fr);gap:16px;margin-bottom:16px;}
    .summary-cards.two-col{grid-template-columns:1fr 3fr;}
    .summary-card{background:var(--surface);border:1px solid var(--rule);border-radius:var(--radius);padding:18px 20px;}
    .summary-card.highlight{background:var(--amber-bg);border-color:#e8c96a;}
    .summary-card.highlight-green{background:var(--green-bg);border-color:#a0d4b8;}
    .summary-card .sc-label{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:8px;}
    .summary-card .sc-value{font-size:22px;font-weight:700;color:var(--navy);line-height:1.1;}
    .summary-card .sc-value.amber{font-size:15px;color:var(--amber);}
    .summary-card .sc-value.green{font-size:15px;color:var(--green);}
    .summary-card .sc-value.red{font-size:15px;color:var(--red);}
    .summary-card .sc-value.meta{font-size:14px;font-weight:500;color:var(--mid);}
    .opinion-box{background:var(--surface);border:1px solid var(--rule);border-radius:var(--radius-lg);padding:28px 32px;}
    .opinion-box h3{font-family:'Playfair Display',serif;font-size:16px;font-weight:600;color:var(--navy);margin-bottom:12px;}
    .opinion-box p{font-size:14px;line-height:1.8;color:var(--mid);}
    .limitations-box{background:var(--surface-alt);border:1px solid var(--rule);border-radius:var(--radius);padding:20px 24px;margin-top:16px;}
    .limitations-box .lim-label{font-size:10px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:8px;}
    .limitations-box p{font-size:14px;line-height:1.8;color:var(--mid);}
    .mapping-list{display:flex;flex-direction:column;gap:16px;margin-top:8px;}
    .mapping-item{background:var(--surface);border:1px solid var(--rule);border-radius:var(--radius);overflow:hidden;}
    .mapping-item-header{display:flex;align-items:flex-start;gap:16px;padding:16px 20px;}
    .feat-num{flex-shrink:0;width:26px;height:26px;border-radius:50%;background:var(--navy);color:#fff;font-size:12px;font-weight:700;display:flex;align-items:center;justify-content:center;margin-top:2px;}
    .mapping-item-header .feat-text{flex:1;font-size:14px;font-weight:600;color:var(--ink);line-height:1.5;}
    .badge{flex-shrink:0;display:inline-block;font-size:11px;font-weight:600;padding:3px 10px;border-radius:2px;}
    .badge-amber{background:var(--amber-bg);color:var(--amber);}
    .badge-green{background:var(--green-bg);color:var(--green);}
    .badge-red{background:var(--red-bg);color:var(--red);}
    .mapping-item-body{border-top:1px solid var(--rule);padding:14px 20px 14px 62px;background:var(--surface-alt);}
    .mapping-item-body p{font-size:13.5px;line-height:1.75;color:var(--mid);margin-bottom:8px;}
    .mapping-item-body p:last-child{margin-bottom:0;}
    .claim-feature-block{background:var(--surface);border:1px solid var(--rule);border-radius:var(--radius-lg);overflow:hidden;margin-bottom:32px;}
    .cfb-header{background:var(--navy);padding:20px 28px;display:flex;align-items:flex-start;gap:16px;}
    .cfb-header .feat-num{background:var(--brand);font-size:13px;width:28px;height:28px;flex-shrink:0;margin-top:1px;}
    .cfb-header .feat-title{font-size:14px;font-weight:500;color:rgba(255,255,255,.9);line-height:1.55;flex:1;font-style:italic;}
    .cfb-verdict{display:flex;gap:10px;padding:16px 28px;background:var(--surface-alt);border-bottom:1px solid var(--rule);}
    .verdict-item{display:flex;align-items:center;gap:8px;}
    .verdict-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;}
    .dot-amber{background:#d4a00a;}.dot-green{background:#1a6b4a;}.dot-red{background:#8a0000;}
    .verdict-item span{font-size:12px;font-weight:600;letter-spacing:.05em;text-transform:uppercase;}
    .verdict-item span.amber{color:var(--amber);}.verdict-item span.green{color:var(--green);}.verdict-item span.red{color:var(--red);}
    .verdict-sep{color:var(--rule);margin:0 4px;}
    .cfb-body{display:grid;grid-template-columns:1fr 1fr;}
    .cfb-col{padding:24px 28px;border-right:1px solid var(--rule);}
    .cfb-col:last-child{border-right:none;}
    .cfb-col h4{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--muted);margin-bottom:14px;padding-bottom:10px;border-bottom:1px solid var(--rule);}
    .cfb-col p{font-size:13.5px;line-height:1.75;color:var(--mid);margin-bottom:10px;}
    .cfb-col p:last-child{margin-bottom:0;}
    .sub-heading{font-size:12px;font-weight:700;letter-spacing:.06em;text-transform:uppercase;color:var(--navy);margin:18px 0 8px;}
    .sub-heading:first-of-type{margin-top:0;}
    .justification-panel{margin:20px 0 0;background:var(--surface);border:1px solid var(--rule);border-left:3px solid var(--brand);border-radius:0 var(--radius) var(--radius) 0;padding:18px 22px;}
    .justification-panel .j-label{font-size:10px;font-weight:700;letter-spacing:.12em;text-transform:uppercase;color:var(--brand);margin-bottom:8px;}
    .justification-panel p{font-size:13.5px;line-height:1.75;color:var(--mid);}
    .excerpt-toggle{width:100%;background:none;border:none;cursor:pointer;display:flex;align-items:center;justify-content:space-between;padding:14px 0;color:var(--navy);font-family:'Source Sans 3',sans-serif;font-size:12px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;transition:color .15s;}
    .excerpt-toggle:hover{color:var(--brand);}
    .excerpt-toggle .toggle-left{display:flex;align-items:center;gap:10px;}
    .excerpt-count{background:var(--navy);color:#fff;font-size:10px;font-weight:700;padding:2px 7px;border-radius:10px;}
    .chevron{width:16px;height:16px;transition:transform .2s;color:var(--muted);}
    .chevron.open{transform:rotate(180deg);}
    .excerpt-body{display:none;}.excerpt-body.open{display:block;}
    .excerpt-item{margin-top:16px;border:1px solid var(--rule);border-radius:var(--radius);overflow:hidden;}
    .excerpt-item-header{background:var(--surface-alt);padding:10px 16px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid var(--rule);}
    .exc-num{font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--navy);}
    .exc-ref{font-size:11px;color:var(--muted);font-family:'Source Code Pro',monospace;}
    .excerpt-item-body{padding:14px 16px;background:#fafafa;font-family:'Source Code Pro',monospace;font-size:12px;line-height:1.65;color:var(--mid);overflow-x:auto;}
    .excerpt-item-body h4{font-family:'Source Sans 3',sans-serif;font-size:11px;font-weight:700;letter-spacing:.07em;text-transform:uppercase;color:var(--navy);margin-bottom:10px;}
    .excerpt-item-body p{margin-bottom:6px;}.excerpt-item-body p:last-child{margin-bottom:0;}
    .exc-subhead{font-weight:600;color:var(--navy);margin-top:10px!important;}
    .exc-indent{padding-left:20px;}
    .exc-list{padding-left:18px;margin:6px 0;}.exc-list li{margin-bottom:4px;}
    .exc-table{width:100%;border-collapse:collapse;font-size:11px;margin:8px 0;}
    .exc-table th,.exc-table td{border:1px solid var(--rule);padding:5px 8px;text-align:left;vertical-align:top;}
    .exc-table th{background:var(--surface-alt);font-weight:600;color:var(--navy);}
    .disclaimer{margin-top:64px;border:1px solid var(--rule);border-radius:var(--radius-lg);overflow:hidden;}
    .disclaimer-header{background:var(--surface-alt);padding:16px 24px;border-bottom:1px solid var(--rule);display:flex;align-items:center;gap:10px;}
    .disclaimer-icon{width:16px;height:16px;color:var(--muted);}
    .disclaimer-header h4{font-size:11px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;color:var(--muted);}
    .disclaimer-body{padding:20px 24px;}
    .disclaimer-body ol{padding-left:18px;display:flex;flex-direction:column;gap:10px;}
    .disclaimer-body li{font-size:12.5px;line-height:1.7;color:var(--muted);}
    .disclaimer-body li strong{font-weight:600;color:var(--mid);}
    .site-footer{text-align:center;padding:32px 0 0;font-size:12px;color:var(--muted);letter-spacing:.08em;}
    @media(max-width:760px){.page-wrap{padding:0 16px 60px;}.summary-cards{grid-template-columns:1fr 1fr;}.summary-cards.two-col{grid-template-columns:1fr;}.identity-grid{grid-template-columns:1fr;}.identity-cell{border-right:none;border-bottom:1px solid var(--rule);}.cfb-body{grid-template-columns:1fr;}.cfb-col{border-right:none;border-bottom:1px solid var(--rule);}}
  </style>
</head>
<body>
  <div class="brand-rule"></div>
  <div class="header-bar">
    <div class="header-bar-inner">
      <div class="logo">${LOGO_SVG}</div>
      <div class="confidential">Confidential</div>
    </div>
  </div>
  <div class="page-wrap">
    <div class="identity">
      <div class="identity-meta">
        <span class="pill">${esc(patentNumber)}</span>
        <span class="pill pill-navy">${esc(standard)}</span>
        <span class="pill pill-navy">${esc(claimNumber)} &middot; ${esc(claimCategory)}</span>
        ${restricted ? '<span class="pill pill-restricted">Restricted Use</span>' : ""}
      </div>
      <h1>${esc(title)}</h1>
      <div class="identity-grid">
        <div class="identity-cell"><div class="label">Patent Number</div><div class="value">${esc(patentNumber)}</div></div>
        <div class="identity-cell"><div class="label">Owner</div><div class="value">${esc(owner)}</div></div>
        <div class="identity-cell"><div class="label">Standard</div><div class="value">${esc(standard)}</div></div>
      </div>
      <div class="claim-block">
        <div class="claim-label">${esc(claimLabel)}</div>
        <p>${esc(claimText)}</p>
      </div>
      ${restricted ? `
      <div class="restricted-notice">
        <div class="restricted-label">&#9888;&nbsp; Restricted Use Notice</div>
        <p>${RESTRICTED_NOTICE_HTML}</p>
      </div>` : ""}
    </div>
    <div class="section">
      <div class="section-heading"><h2>Executive Summary</h2></div>
      <div class="summary-cards">
        <div class="summary-card"><div class="sc-label">Claim Number</div><div class="sc-value">${esc(claimNumber)}</div></div>
        <div class="summary-card"><div class="sc-label">Claim Category</div><div class="sc-value">${esc(claimCategory)}</div></div>
        <div class="summary-card"><div class="sc-label">Percentage Mapped</div><div class="sc-value">${esc(pctMapped)}</div></div>
        <div class="summary-card ${ec.card}"><div class="sc-label">Essentiality Decision</div><div class="sc-value ${ec.value}">${esc(essDecision)}</div></div>
      </div>
      <div class="summary-cards two-col">
        <div class="summary-card"><div class="sc-label">Weighted Mapping</div><div class="sc-value">${esc(pctWeighted)}</div></div>
        <div class="summary-card"><div class="sc-label">Limitations</div><div class="sc-value meta">${esc(limLabel)}</div></div>
      </div>
      <div class="opinion-box">
        <h3>Opinion</h3>
        <p>${esc(opinion)}</p>
        <div class="limitations-box">
          <div class="lim-label">Limitations Detail</div>
          <p>${esc(limBody)}</p>
        </div>
      </div>
      <div style="margin-top:32px;">
        <div class="section-heading" style="margin-top:0;"><h2>Mapping Summary</h2></div>
        <div class="mapping-list">${buildMappingSummaryHtml(mappingItems)}</div>
      </div>
    </div>
    <div class="section">
      <div class="section-heading"><h2>Claim Chart</h2></div>
      ${buildClaimChartHtml(charts)}
    </div>
    <div class="disclaimer">
      <div class="disclaimer-header">
        <svg class="disclaimer-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
        <h4>Disclaimer</h4>
      </div>
      <div class="disclaimer-body">
        <ol>
          <li><strong>Preliminary and Informational Nature:</strong> The present work product was generated using a prototype AI model and is provided for informational purposes only. It does not constitute a legal or technical opinion regarding the essentiality or non-essentiality of any patent claim to any technical standard. It is not a substitute for legal or technical advice, and clients are strongly encouraged to seek independent professional counsel before relying on this material for purposes such as licensing, enforcement, or infringement analysis.</li>
          <li><strong>Scope of Analysis:</strong> The analysis is limited to the individual patent claim(s) identified in the chart and does not take into account the full patent specification, including the description and drawings. Consequently, any interpretation of claim scope is based on the claim language alone and may differ from that reached through a full legal construction under applicable law.</li>
          <li><strong>Referencing of Standards:</strong> Where citations to section numbers, table numbers, or figure numbers in a technical standard are provided, they are included for convenience only. While care is taken in referencing, these citations should not be relied upon as authoritative without verification against the official version of the standard.</li>
          <li><strong>Interpretation of Standards:</strong> References to technical standards are based on publicly available documents. Where relevant, excerpts are cited in text form. Figures and diagrams from such standards are not reproduced; instead, any associated visual content is paraphrased using descriptive language. Such paraphrasing should not be construed as a verbatim or authoritative interpretation of the standard itself.</li>
          <li><strong>Subjectivity of Essentiality:</strong> Determinations of potential alignment between a patent claim and a standard may depend on how specific terms or functional steps are construed. What may appear to correspond closely under one interpretation may be viewed as merely analogous under another. This assessment is inherently interpretive and does not reflect a consensus view or judicial determination.</li>
          <li><strong>Implementation Considerations:</strong> The presence of a feature in a standard does not imply that all compliant implementations necessarily use that feature. A compliant product may omit or bypass specific technical elements referenced in a patent claim.</li>
          <li><strong>Alternative Solutions:</strong> Standards may include multiple options or alternative techniques to achieve similar functionality. A given patent claim may correspond to one such option, but not to others that are also compliant with the standard.</li>
          <li><strong>Legal Proceedings:</strong> In the context of litigation, essentiality determinations typically require a far more detailed analysis, including expert testimony, claim construction under applicable law, and examination of implementation evidence. The present assessment should not be relied upon for litigation, licensing negotiation, or investment decisions without further professional review.</li>
        </ol>
      </div>
    </div>
    <div class="site-footer">ipmind.ai</div>
  </div>
  <script>
    function toggleExcerpts(btn) {
      const body = btn.nextElementSibling;
      const chevron = btn.querySelector('.chevron');
      const isOpen = body.classList.contains('open');
      body.classList.toggle('open', !isOpen);
      chevron.classList.toggle('open', !isOpen);
      btn.setAttribute('aria-expanded', String(!isOpen));
    }
  </script>
</body>
</html>`;
}

// HTML endpoint — POST the IPMIND analysis JSON, returns HTML string in JSON
app.post("/generate-html", (req, res) => {
  try {
    const body = req.body;
    const data = Array.isArray(body) ? body[0] : body;
    const meta = {
      Patent_Number: req.query.patent   || "",
      Title:         req.query.title    || "",
      Owner:         req.query.owner    || "",
      Standard:      req.query.standard || "",
    };
    const html = buildHtml(data, meta, req.query.restricted === "true" || data.Restricted_Use === true);
    const safeName = (data.Patent_Number || meta.Patent_Number || "report")
      .replace(/[^A-Za-z0-9_-]/g, "_");
    res.json({
      html,
      filename: safeName + "_report.html",
      patent:   data.Patent_Number || meta.Patent_Number || "",
      claim:    data.Claim_Number  || "",
      features: (data.Claim_Charts || []).length,
      excerpts_total: (data.Claim_Charts || []).reduce((s, c) => s + (c.Cited_Excerpts || []).length, 0),
    });
  } catch (err) {
    console.error("Error generating html:", err);
    res.status(500).json({ error: err.message });
  }
});

// ═════════════════════════════════════════════════════════════════════════
// START
// ═════════════════════════════════════════════════════════════════════════

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("IPMIND docx service running on port " + PORT));
