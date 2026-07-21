import { useState, useEffect, useCallback, useRef } from "react";
import * as XLSX from "xlsx";

// ═══════════════════════════════════════════════════════
// GHAFES LOGO (embedded)
// ═══════════════════════════════════════════════════════
const LOGO = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAFAAAABQCAYAAACOEfKtAAAgCklEQVR42u18eXRcxZ3u96u6t293a98tyfK+YRsvYLMbyyQEQliGEJkZ2yEBTpLJvBNyTuZlm0xGdt6beXnZZpgkQ8KEEIIdEmlYDD6EvAxIgFnMYmy8gWVb1i5bUqvVUm/3VtXv/XG7ZRkMtmQ5JJOpQx+Z7rtUfff7bV9VXeCPqDEzMbNoamKL+bQfycyCmemD7LP1QQPWCIgygJoBQ0QGAAMw47yOaAZELWAIYBDxf2kAswMmIgVAj/k+Nw1MSyjMdBVqXK0rDKjAAEEwIAUSkjhqkewNWmgPWjjqAO1ElBoLehOzVXvigZzT9gelfwOzrAM4OzBmtmOetzIJ+6q4Z1YlPbM4pVApQhYpASj2P1mEJQCL/I/UAKc8HbSoM2yJN3MC4rkcG8/kAG9QhoHvvN+fLIDMLDYC2JQZyDDzkrjCXw2nzV+MGLEgbQEDKSCSAobSQEJDpzQ45Sl4WsMY3yKlELAtgaBlISghwhZEvgOUBIFiB7DTBvkB7M6xxSPFNn4dJDqYvT8AnAsg6Vz7OPimqgFg0OOrRgy+GEmqjyUCluyIAb1xcNSD7o/FaaA/ImKRQaRiMeJUAkJ5EGwg4ANoQNBCwkgbMhxGML+AC0qLuaS4kIvzwlxkw6rMAabmAWFPpwod2lrgiLsLiF7K9EfCN23+owewoYHl2rU+cFGXVwwz6qMK1/d6wOEIo1+R6uofEt1tnSLe241cL4GaPBuzy/Ixo7wAlSV5KMoNIScYgCV9CD1lEE+lEYkl0B0ZxtFjQzjcN4zOuEIymIf8yipUTqs2lcV5ptxma3YxoZQMioOioZDxrTyH9mUZOVlsPCcANjFba4gUM+f0umZTxMMXe5Ww3urXptcjPtrWI3pbWqgoPYSVNYW4/LwaLJw5BQW54QndbyA6gj2tPXjhQCd2dscwEi5G5fy5PL16ipliK7Gw3KJyaZJFFr47pUX8Iy0mt6mJrTVrSP3RAcjMkoj0QIIvjQnz7xESi3Z2euhWtj7c1iN79u3FfMfDjRfOwhVLZiEUDIw9F4Y50zFC5j+ACOx5INsGMwMMcMasBQnQmFEMx5N4dtdhPPHGUbSaMKoXL8asqeW62lbygqkWihiv5nvuZ4pznd3ZB/1HAWDG3xERmWNJ/vwgm7sPjQj7rahRR6NJ6+BrOzGHY7jtyoW4ePHM0fO0YWROPAkI/6IGIIHkjuegowMgJ4Sc1ddkek4nAZ/BFVKc+P7ZN1qwefvb6HBKsODCZTwt39GLSoQ1M4x4Huu/rgpbm8/WL1qTDV5XUn9/MIAv7WhT3OYGzN632yzVshdfuGQmrr981RimAULQSQN+r+fLqSRy1nwUyZef80EV8uSjxoDPDBhmCEFYvXwurlw2Bw8378bmpqfp+KJlVoKn6n7Hy7mw2n6wO8Ezieh/NfhVzYRAtCYDvEZArCXSHQn982hA3L691VVHU1Lu2vGqWEZR/O0dtSgpyvNZkhmcpFNe7NT/LyVgjP/Xv8gpsKbRPzLzb2MYJAifWLMMly2eju89tgM7+vtlasUSTrSlzapZzrc6krqwhuhvm5qaLGbW9AesYsDM1MRsAUBHXN+3VzP/uMV1v7Qrzdf+sIkf/N2rnG1Ka55oS7zYxCYZ5/hLTRM6f+y9f7L1Rb72nuf5K3s8c88h191vmDvi+juZ8Vh/aAbKNUSqfVh/ezgk7njuiOe1DBv70Pbn8dXLp6H2gnkw7Ps4KcT7X0lrqIE+vxQeyzBmqEgfAIIZikJ1twOWfeIYIoAZsqgEFHBO3UkhfD8J4HM3XoqpL+7Fvz27ncyqy23R6nlXzrS/3D6kjhPR95jZonEEFjrbaNs5pD49kiPv397meQeG2W7Zvh3/cNUsXHjeDChtYMnTAGcMIATS+3dBHe+FVVEJGM6GX8BowLbhzFsMt+0QzHAMZFknAahjURADoUtWj17vvZoyBpYQaN55EN9/pRcLr7iMF+SwvqTKtvLT6oapBfa27NjOGQMziageGObFx8ncs7PL04eTwmp55WV8o3bmmYM39pquCxEKQ+QXZaJv5tkahgiFASLIolKQtDJBhEc5wIZhIn1nZnJCQGmD2gvmQWuDu1/ZSYHLLhI5fR4vKZIPJBK8DEDnmSbbYoJBg5jZ6tf6gfaUCB5J2di/ax99fkkxLlo0c3zgZaJwYN4iiLx8mOEozHAMJjbkf4ajiD///8BuCokXnoGJDsAMD435fQjEDGfxBe9Kb94TROmD+KGVC7B+poN9e94ShxLS9ChR3J4292UCCZ0rBoq1RLo1qr6WDMoL9ncYtf9Ij/WRwjSuvXiJbyJyPM/F76cI58BZuOzUVh6LAsZAFpfAOX/FaS53Zl7JkgLaMG5dsxRHHn4JBzr7ZS6VqJJp8ur2Ib6diO4/E1MW4zVdAGYgwTXDTN/Y26vN4WhKlgy04TNXL4VhhiQxUa/q+6+xH60BY8CeNxpooLx3H2fMqVOb05KfYBi469qlCHQextG4Fgf6jIkq/W1mLtq4EXw6xXu8oyUi4oGE+WZMityjKTLHjxyh/3HZTNiBwHgIcGomCnHqz9iLnskxZzwYP8rn5ITx2RVV6DzcKloTMAlHlrdEzP/ctIlMczPkpACYDRyDSZ45wnzbwQFjDh+LyisLNBbMrIQ2DEEf6PTEhJoQBG0YFy2cjuWBOI70D8u3+w2PaPxNLMZla2qh34+FZwxgc7N/7EDc/E3SkU5Hkg339dCtK2eAT8SCP8lG5Mf09SumIXmshzqS0OmQKDzu4Q4QcTPem4XiDNlHa9aQYubcIZc3HI0y2geGxeUlAqXFBWBmEP3pIiiIwMyYVlmKC3I12gfjoiMGHkqrO5nZrgU0cGoWijNknwSAowPqIzosp3QmWOtoRFy7oAKM/zqNAVw3vxTJwYjoGDGsHGtuZxSXExEznxqrMwKwr9bHKaFxy5AGd8ddnut4mFpRBDD+ML6P+X0rjMlgIQDMqy7FNJlCd8IzcQGOpc0tAND8HnmhOBPzXUukmTkU97C6Nw4aio2Iiyr96sDwueegCOcgtfs1uAf3+1XIObqnMQySEivLgxiMxcXxOCih+EPMLNcQ9IQAbGz0j+nox2IlZXVvXHPATYollQVjC4lz2vRwDBRwELzwUnjtR0YFhHMRTABgaWUeKJ0Ux+IGaeZ5XX2Y7RdfLMYNYFmdT92UxoXsECKu1lMCGuWF4VEx8xzRAWCGjkbgtbaAPRfBxRdAHe8Bu+45NeOpxbkokQr9aa0pbMkRrZePzUQmlMYkPbM0aYDhtIcZORKQFjRPso8zxhcSxiTM6ngPwquuhlVV48tW+UUnJ86j501OZwwDViCAmjAh5ip2AbiGzgcA1E6gFq7NLJnwmOeMeIDrKqopDYyJWzR59pMBhd00VN8xqEg/3EMHkH/jX4Kkn4qlWw9CDQ/BqqiCVVYB4QT9SadJ6glnrlSTa2PviEJcAQXGzM1gweMGkIgM+TpaZdyTYKOpPMc+WzHRH7RW4ETCnwxKxWEiffAG+mA8BZmXh1BZBUJXrEF654uwZi2A6u1Czuy5EMEQVF830kfeAqSEXVwKq7QcCOf5FVEoBMqKruN0MdmjK3JsmCFNCQUoTVUAsHHjxvEByMxERGyYgy8eVQVJD7DAKMxMRdJZeut401MQoRAsx8GhIQ9RKwcong0qLIEVDMImRtgmlM3NhWk9AGtKNfoLqzCYMkiFpsFLpeANDgDHjkEe3odpeRJFAYI7PIy8j9x4NoU5CoIWyBhKKUAxF2WwMBOSswYH4RjDQdcAEkwhW0xM0M74KR3pR+Sn30Xe9WsRXLICh6MeWiNpBCxCgA3gekinXSjDSLgaeko+5i2/GAbAW4cGkPQ0glJASgHKK4IpLEfMM0iQwZJpecg9vA993/kGCm+9A/a0WePKIbMjCtsSkgw8DRjmMPxyTmVJNa4gQgRiELEBLEHj1PveEVmJENv6EEwiAWfhMiQ9jdaeQYR1CgEv5ctVRCAhYEmBoG3BAoONARsDx5ZwLAkh/UDCSoHSCeSaFLxkEu3dMQRmLwCkROyRByfAQhrVCwX5QSUzrTPxUq5Qw3B2/R3RxIw3wwI90IfkjmeRe81fgCwLkREXhgRICBAJENFJuaXOTAb5v/vKib9Yi0bVFCIBJgHbtjCUdKGERP6HPoZ0ywG4rQf9PhszIVdDANgXVc3E1ZgSpAVzyhaAIcGe4Ymz77EtkGVTEFy8HCll0BlNIZCZJDbMcLVByjNIehopz8DTxl/OMSbNcHX2d/8YZRjMDEmAMoy2SArWzLkIzJ6P2MNZFvIZRmG/eQz/oQiADOJZZfqd88biNBGYMwlmmpgGgxLQICTHmwBm2RfpR+KlZoQuuwqyqATHh10MphQ8wxhJawynfUB8ZZsQCkjkBy0EbTlqXnmORK5jIWj5jFSGEXf9c0fSGgxCW38cOqcAuRevQvrtfXCPtgAkxsXCuGIwCXYEIIgHAKB+IpVIQwNLBiCJusMWwEJw1DUnPa0zZt/Dv4QsKUdo0VK4BuiLJRGQAoIItiQEJCGQAYYBKG3gaYYeZSBDGYanzeh3liA4UiAg/WsIIhAMukc07BlzfBb++ufjYiEARF0GhOCQBQhQBwDUTqQSyZZyZLglxwIEST6e4vGzLxpB/PnfI3TJatiV1RiIe+iKuXC1b6bK+ODozMcwQ7PvA7MrVBkY/V0b/3tlfICz57vaIKkYh/tGYIrKkLPycqT278rU0GfOwmNphk0CIQFIprffqxI5vQ9s9v/YoF0OAyEp0JmcgO9ruB+ytALh+YswDAe7OwYhBYGIYAtCUBLCFiFoERxLwBIEYwxSroKXdRkMpFwNT2n/PEvAkYSgBELSZ6IlfBbGky52RwGrehoC8xYh+uBPzkiE8IM2oysJ5EhB0gME85sTrkRqa/3okyPla5G45kIpZWfKgzEaQsgzYp8ZHsLI09uQX/dpyLIKxFIKpSELli1hSICFBZORqZSb9sUCQQiGchBkwLFOJBh5BbmQrgKlU37dHAhABgIQQoCMAmkFwQbEFgwB6dxi5F58Jfrv/T68rnbYVTWjy+ZOFUAIQMrT6EkTFwcsqUZUKt9yd/mVyIRKOf+k+Um81S24tSJsz3pzKG164lpU58n3r5YyK6qGfn0frLIpCMw5D6KsEkEPsPLy4JBBwE3AHu6DNTQAkRgBGe2LCErBdT2ki6Ygd9lyPxkThLz2Awj3dCAUDMC2LX/9tJBAOBemsAReXjHcQBhpFmBjECyZChM9DmfBEkTv/1eU/f33AG1OOcuRHUvbiMawEWZhjpSBtNozf15O51lUIsRNTWzRYnJfOOA9PSWMmbuHpNkb1aI6733khCz74iMYfvzXyF97uy9P9R9DTjSKBak4LOX6Tj8UBlWXA6EZJwRTSwJ2AHy0BenXnwEu+xCSLz6DmfkFkNdcDXgu4PlJN7QGpxLgkRi4tw9gBgccqEAIxIVg10X4oisQ+fG3oY51wyqvfE8WAsCeqIaE4FIHcFzxFJ+Y1lATKuVqM5J+iNGY8vCZEinEriGFj0x9n8mkLPsa7oc9bRbyrq+D13oIOhrxdTcnBM8YX3XpaIV7tAWqsw060geTTIIsC7KiCkV3fBGBGfMQ3fwT5Ky+Bm50EJEv3gbV1Q72XIhgEKKwBPbU6bBnzIFdMwuyrAJCBiAAmKFBMBjhVVcj+doLGLz/X1H2tW+fkoWCAE8b7IkxyixLihHNISEfGevK3qv0O10oJTDA+2A/DbW/w7JmvR5L8F3zbDG3wPZXm9I7bQHQ0UF03HwZZGkFcj98g7+SauA4vK52eO2tUL2dMLEhsFa+XGXbmcVDPjM4nQaUh+otv4dVWg5WHjo+sRqcToJC4dHMmrUCPBesNSAlZF4+rIpqWDUzYFdPh1VWAZBA4qVmpPe8jpqtL8Gunn4SC7NjeKPfxX2tSl9WFBZTkuq1q863Lt64EbRpE50NgEB2VfuLe92vD+fb//Sfx1JqTgGsz84NvtsPZiKde6wbiVdfAMAwifgJBVtKkO2AbBtk2ycmZpl9PY59l07Sgh4ahLvzZVhVNVA9HbAXLoNVUQVWHvCOhegAwIYBo8GeB3bTgFbgTCYggiFACATPvxDBGXNOkrt8V8T45wMp9MWh15SHZM6IvvPyRdbP329FP515OsdEAA62oLTVVQcPMwoOJF186bwAzci13s3CSZxq1KkU2EuDLBsyFJ70rQXZvu+LKtzT4prluQ7VKNNVYewFS5cicaoSblw+MHuBpia25q2hvu27vR/OLrK+2ZZS6vEubd0133pXKNODA1DHugHlQUcHfTPLSu9SAmzAqRSCS1dAllaATUY0eKdPJYK0AyDHgTEM9jyYDKP8pS0MAoOEgOruQPrAblAoJ6NEaJ9hQvjgFxUDJGBXTYMoKDyJgYYZT3QrFJIw03KkFRo2/3fZEoqfbj/JuB4mM9PGjaDbbkPekYQ60GlRxWsxFxtm2eKS0gwLM6mS29qC2CMPIrhoGUIXXekP2LIAaUH1dCD5ynaw9iBCOZDFZXAsX99T2vis0yajtPjqjKcMnICEpwxsS0IbAzD8c5SGpxmqr8ffTyIthFd9GLKo1DdhpcBaI/HC0/BaW1Bw6x2wKqf6wMG//tO9Hh5p98ylBUGqcHXL1UF7yca58Db6mxV5UgAE/B2Qa4n0jr1q3Uiu3LIjklaDxNbXFjooDrx7biL5ynNQx3qQd8Oto9+l9u6ECIYQmHPe6PFvdw7hjUPHcdGCCijNKC0IYiTpQWlGPOVhalkunn69A/OnFaG7fwRCEGZOKcCe1n4snlGCudUFJ66/+1WIvHwEZs0f/S72H79EYO55CC5decJ0M6VYT9LguwfSqBJSLc8PWHlJ9dGLF9tPjd2u9p4zeeMFcC2RbmCWFy+2fiVi3uPnFzkWKagHWr2s+DgqP+lIP3RsCPaMOeBUEmZkGCY+DE7ET8ymGY1kykPzzjbMq8rH/iP9+N2OVtz/5F4cbI+g6fU2bHvhMA53RLDzYC9Ya7zwZhfSKQ/90Tj2t/bjSPegv/fE80ZrXf9eI/79kgnYM+f6KdLwUNaaQAA8w7i/1YNjoBYVBywR935x8WL7qQY+PXjj8oFj276N4Pp6FjVhfKY9pVcuzrUrXx1Km9+0k/ir6b4pq86jGNpyLwJzFsCEwki8+IzPNcuC13YYObXX+r4JgB2UKCjOx97uOIZG0iivKEJ1SQ72t0cgAw5Kyxy8fjSGqy6ag9kzyuC81g0KBiGDQZSXF6Ez6vm+085Mdlk2Um++Ch2N+Ao3COQEoXq6MPjvP0DBJz8PUVoBAcYvj3o4PmL0xYVByxlRh8637buyC0nHMwUw7pal9ytvuqtjjnjmkGv4zYQnrptq08eqbKhEHEJaEIGAH0vHZP3p/bugI/0IzF4ANgZCCiTSCoe7hjCzMh+GGblBG90DcVQUhWFbAh3HRzClOAzHljjSM4S0p1FRFEbPQBzVpbnICwfAxoAEIbVvN5y558GePvtdooZJp/w4Fgqhod3Dc73KLM+xMdMS6WBKXX75MueN8ezmPKuMYDQ33KPuTOfJn+0dctVbKS2vq7HpukprzDp6nNhhRIBJJpB68/UTGwmZIQXBCVhIu9qvzgwjYAm4ylekHdsPIIYZwYAEEcFTZvQYbYyfyRkDsm0El6zwc8xsD94xd9zY4aG5W/HisDTzwwHpjKQ/ccWy4MPj2eJw1gCOBfGFN92vugX2t/dGfRCvrLRobY19Up71QbZsHzQzNh9VeLVPmUUhyfNzAzIw4n1u1ZLAvRPZAjspwxoD4tfcPPv/vBVzzb6UxoIiKT45w0a+7S/mJhrdPgiYc7yqK1sOZnUNAvrTjAeOemgf0npx2JKzQjachPfXq5YEfjrR/cOTxosxIH426ch7ujSJPSMp5TjCqptmYWmhHGXCCSDPXRsLHADsGNB4pMMDPFZLC0JWBUzSSatPr1riNJzN5utJHUe2I6+86X044dADQwFZtSeSUr3ayAtKbfpYlYVyh0aBzCogk22qY6/blTR4okth76A21ZJ4SUlQ5qV0SzhpNqxcHnjlbHeuTzoRsh16+eX4VF0U/HHKETceiRu8HU8rWJCXlFm0usxCqUPvGvREmMkntIuTHkZvyqDpuMarfYotA31ebsiaFgRCabO5YHD4i4svK4xMxrb/c2JJ2WoFAF7ep+5MSHwrHZJVB6MKR9OekhJycZGki0sk5uRJWPRuLfZ0HvJUYHsGeHtYY8eAxv6oMdBsZgUD1pwCCSdpDueBv75igdWYHfhkeOFz5oqYmTYCtInIHDzIZQPGfDnB/Nl0SBYcHWG0JVI6SUBpkMS8fEnz8wRqwgLFATrj1RiGgYjLaE8YvB0zOBjTHEmxCYMwIycop+cCgYTuyyH5o+Io/mXeJRTL5nifenSw8IGbi6J1DQ2ysa7OTPR1UecEwHpm0dzcLMr7arluIeTaxeQCwL59PD0mzecSzBt0UNZEDNAxrNHnejoFZscCFQaIShyiYkdQgQ2EJcHK5ODKAHHNiLnAgGs4kmYedJldBXYYYkowIGpyJQoJsNL6SA7hF3km9bOFC3N7AKC+YW9gU90ib8MT3V8As7f5pqn3nCQaY/wg0rkAb9Mpsvifvsb251aQBwCHX4sUDOUVXRc3Zm2KzZXGsYpTEoi4QCTFiLkeksYYBeaxa0/JL97JYoiwJSnfDqA4CBQHAEcDlFL9IUnNOUL+pjKB31Yuo3jGGuTaRqBxLekNj3fezoz4lpumNqxuarLmmhVXDg0kXmlcWzECZhovEycVwPp6Fps2kVn/ZPdKi5xaIgitvYhR3p4tN9W8PDbIZM85dIjLhzx1SZqxymVe4TLmKqBc2LYN2+cFZ/0dZzjiAcb1PJtwzBbiYIDo1ZAQz4cYOxYsoP5RX7yXA3WL4GUqHl63rW0WeeLqLTfX/HTDtu5rYfAVJzcf6eEo0mXmYw2X1qR893HmINJkM2/9Y51fIkE3kLS/zyr1lm2HA56XvpSBGYZw969uqBrY2NwsN9bWMk7xYjDu4vDeOCqVwZS0VmXEVKCNCbIQTEBaEA85kvvCZPfMy0UvVVPipH40sbWoDCLrNgCgroFl41rSGx7tvIGkuJzZMAmxUnnq7x76+LRX1m3tuI9Av9ty09SG7LFnOm5rMsG77Ymei4zRn5Z57kUPrKlOjTlk/6e2dV8ISCsjTqpNY4INANG4D3LfIigiSgA4nPmcttU1sKxbCIlF0I2NjdiUYXfd7yMFTmLk0nQ69kxjHTwwEx7sfYYL+SqAjj94Q9WHs+cTdf+GDFd/YEFkdRNbz64htWFrx/cB0br5puofXXv3Qeepu+a6p/IpG7Z2fJ2N2O56VS8CQONamKzZZFXvRYtAZWWgvjKIsr6TpaXmMohF+6HfodcRAP7k7yPT9MjIFSRkCUCDxfZI4w8/eop+jPF36x/v/IahQMNDN5S3vHMF6h+EgSfimOgjrYMAMC0419Q1QthbOz8KY9rm7qrZBwD7FzUSA9drkfr5WFP55GOdywXntBJRNDNHdlo56QtPsjMoo1WUjk958MapL23Y2rUJaXe5FNYvWGAIykzpHSmwAbg+wIzVTZCVQ+3lvybqvv3J7jJP6Q3EOLjlxvKW+vrxv5RsUgB8thYazIRH2h6AJf/3uq2d1997E23z2dZVwULcuWkT3QwA6x5tXSaFHYAWf/np3w50s5fa44RCPYlE4m4RSN4yWl6AqX4jqHVF74qZ1095LRvZ73jsQJ4rC/6OhN0WNT1xWwaLXEYvgJekpMe19j4BokpSdIjD8uHGayriWcZ99nVY964gb8Pj3bfe/sxwZXo41i+Jnn/wpqkvZQPguDWLSaFehvJbbpnRg6B1FxGmbdja9dVPbjteB/BagvlyfX3mJYjCukV77g/YUr80ZPoV5BWu4R8zY/f911X11TWwXN3cLEHEhy7oXGeM3h75LWzfXzVIPRj02PCV0pYWMY9kKpcImOmB66teB/AUmL0Hb67+3ewXpyTr61nUNTRIEPG9mTRq841V/2xr7x9/dVP1dx68ceLgTa4JEzGYaTNRHMC/rdvWViQg64zRP//VzdMP1TWwBDPR410LtCXufuj66kEATQCa1j/avlAI+SwAOl4Gqm2uNZWPtVUJolls8JuI1/1xAA+F4yvtB26fmVr/WMcTRutling7C7yee6y3B6hGfT2LI7rrn4zgzRt+uWvL/mBjqnHtWg0A67Z1TxdaX5NOT72vcS3MvVfTENhn+UTBm3Qf6DORqa4B4lfX0yCAe7OpxaY1pDY81rWewfqhm6b21zextb+vkdOB88NEYgYHxTezxrtpE5l1j3V8wljBfyFvpARGbATwkHP+jIzPtPZ4yYT70Men7xq99+eA5qYm+eyaNQPrH+s4hILSusablv3CzwxMLYwpYKJnGutgMtI0gYg3nWVJfA7e4kvcuNbf4V3XAAE0AtmFOcT5RpsfAUw9eaDGNWvV+se7LgWjdfM1lfFrnzzoPLWG0hu2dn1FBpzVnpdsIci0sOT8ddu6p9+7gtoyQsJ+Ia0N76x8amtrzbPMhMc7twDYsmFrV4U2Kg3W27bcNGP/uShbz+FrkLNAnmhja897n2ANZqKtnXeyxD1gpqeI0hu2ts9hNvPgWP8gXBQAMgAS+0mpWwD8oK6hQYarp3QmurqcTz06WLjJj9oEEG8iv1ZxGw4fCDjBvw/Y9J/3X1fT905xA3+qrb6eRTaYZCPjrY+0za6vrxdgpg2/68lZ/1jnz9b9R9usseetf7I/f/3Wjp3rn+zPz56/fmtnw7pHO5dlr/t+iTY+4Led/8FaXQMHbn+yuyxrmn4EZQkQNmzt/Pz6h49WZoH/1LZjU1Y3Nb2XBdGfDXD19fWnWKz8roH/eTBokpRYek9THPvqkT8Xs/zv9t/tfdv/By5r8fhC34tyAAAAAElFTkSuQmCC";

// ═══════════════════════════════════════════════════════
// CONSTANTS
// ═══════════════════════════════════════════════════════
const FELLOWSHIPS = [
  "UCF Legon","GCF Legon","CMF Korlebu","AHCF Korlebu","CMF AHCF",
  "CMF ACM","CMF UHAS","AHCF UHAS","GHAFES UHAS","HCF Korlebu",
  "GHAFES PU","GHAFES UPSA","GHAFES Wisconsin","GHAFES UCF City Campus",
  "LCF GSL","GHAFES RMUCF","PASAG CF UHAS","GHAFES HTUCF","GHAFES VVU","CMF FH",
  "GHAFES TTS","GHAFES UHAS Hohoe","GHAFES KTUCF","GHAFES Ashesi",
  "GHAFES UCF UESD","GHAFES ATUCF","GHAFES GhanaCU"
];
const MONTHS  = ["January","February","March","April","May","June","July","August","September","October","November","December"];
const YEARS   = ["2024","2025","2026","2027"];
const WEEKS   = ["1","2","3","4","5-Transition"];
const DEVLIFE = ["Struggling","Fair","Very Good","Excellent"];
const LEVELS  = ["National","Zonal","Local"];

const ZONES = {
  "South East Zone": [
    { name:"Sub-Zone 1 (Korle-Bu)", fellowships:["CMF Korlebu","AHCF Korlebu","HCF Korlebu","CMF FH","CMF ACM"] },
    { name:"Sub-Zone 2 (Legon)",    fellowships:["UCF Legon","GCF Legon","CMF AHCF","LCF GSL"] },
    { name:"Sub-Zone 3 (UHAS)",     fellowships:["CMF UHAS","AHCF UHAS","GHAFES UHAS","PASAG CF UHAS"] },
    { name:"Sub-Zone 4 (Satellite)",fellowships:["GHAFES PU","GHAFES UPSA","GHAFES Wisconsin","GHAFES RMUCF","GHAFES UCF City Campus"] },
    { name:"Sub-Zone 5 (Others)",   fellowships:["GHAFES HTUCF","GHAFES VVU"] },
  ]
};

// ═══════════════════════════════════════════════════════
// DEFAULT STRUCTURES
// ═══════════════════════════════════════════════════════
const mkWeek = () => ({
  det: { membership:"", attendance:"", submittedPrev:"YES", whyNot:"", whenExpect:"" },
  wit: { awareness:"", exposure:"", converts:"", approach:"", trainProg:"None", numTrained:"0", trainTopics:"None" },
  cdf: { existing:"", added:"", totalMbr:"", metWeek:"", present:"", leadTrained:"NO", trainTopics:"None", resourcePersons:"None" },
  pry: { engagements:"", inPrayer:"", studsTrained:"", trainTopics:"", dailyQT:"", sgbs:"", sgbsTopics:"", memorization:"" }
});
const mkReport = () => ({
  hdr:  { studentName:"", fellowship:"", presidentName:"", specialProg:"",
          month: MONTHS[new Date().getMonth()], year: String(new Date().getFullYear()),
          subDate: new Date().toISOString().split("T")[0] },
  weeks:{ "1":mkWeek(),"2":mkWeek(),"3":mkWeek(),"4":mkWeek(),"5-Transition":mkWeek() },
  lead: { fresh:"NO", final:"NO", llts:"NO", leadBlock:"NO", integrity:"NO" },
  zonalProg:"NO", majorProg:"NO", programs:[],
  other:{ staffVisits:"", assocRP:"", mspc:"", alumni:"", quota:"", discipled:"" }
});
const mkCMC = () => ({
  hdr:  { staffName:"", station:"", books:"", devLife:"Very Good",
          month: MONTHS[new Date().getMonth()], year: String(new Date().getFullYear()),
          subDate: new Date().toISOString().split("T")[0] },
  lead: { fresh:"NO", final:"NO", llts:"NO", leadBlock:"NO", integrity:"NO" },
  zonalProg:"NO", majorProg:"NO", programs:[],
  other:{ fellowVisits:"", assocVisits:"", assocCalls:"", assocRP:"", assocMapped:"", discipled:"" }
});

// ═══════════════════════════════════════════════════════
// STORAGE  (uses browser localStorage — works on any device)
// ═══════════════════════════════════════════════════════
const STORE_KEY     = "ghafes_ministry_v2";
const SETTINGS_KEY  = "ghafes_settings_v1";
const UNLOCK_KEY     = "ghafes_unlocked_v1"; // sessionStorage — resets when the browser tab closes

// ─── Staff-access gate ──────────────────────────────────────────
// CMC Report, Zonal Report, and Google Sheets Settings are hidden from
// everyone by default. Each is unlocked separately with its own password.
// Passwords are never stored in plaintext — only their SHA-256 hash is
// kept here, and the value typed in is hashed and compared to it.
//
// ⚠️ CHANGE THESE BEFORE GOING LIVE. To generate a new hash for a new
// password, open any browser console and run:
//   crypto.subtle.digest("SHA-256", new TextEncoder().encode("your-password"))
//     .then(b=>console.log(Array.from(new Uint8Array(b)).map(x=>x.toString(16).padStart(2,"0")).join("")))
// then paste the printed hash below in place of the matching value.
//
// Placeholder passwords (CHANGE THESE): cmc-2026 / zonal-2026 / admin-2026
const ACCESS_HASHES = {
  cmc:    "da00da165ee8f23f3c05da976c21dfdd09d424ed19743e856b5a902598364209",
  zonal:  "619474a9236b8743cc27d1d67d8f8b4baaac3421a0ad97ee02c9a7ac8a06b2b3",
  sheets: "dea0eb2bfaf38042753851289edacfa858ce5117d17bc8bb6814dbce7119daaa",
};
async function sha256Hex(text){
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(buf)).map(b=>b.toString(16).padStart(2,"0")).join("");
}
async function checkAccess(role,password){
  if(!password) return false;
  const hash = await sha256Hex(password);
  return hash === ACCESS_HASHES[role];
}
function loadUnlocked(){ try{ return JSON.parse(sessionStorage.getItem(UNLOCK_KEY)||"{}"); }catch{ return {}; } }
function saveUnlocked(u){ try{ sessionStorage.setItem(UNLOCK_KEY,JSON.stringify(u)); }catch{} }
const rKey = (f,m,y) => `${f}||${m}||${y}`;

async function loadData()    { try { const r = localStorage.getItem(STORE_KEY);    return r ? JSON.parse(r) : {}; } catch { return {}; } }
async function saveData(d)   { try { localStorage.setItem(STORE_KEY,    JSON.stringify(d)); } catch {} }
async function loadSettings(){ try { const r = localStorage.getItem(SETTINGS_KEY); return r ? JSON.parse(r) : {}; } catch { return {}; } }
async function saveSettings(s){ try { localStorage.setItem(SETTINGS_KEY, JSON.stringify(s)); } catch {} }

// ═══════════════════════════════════════════════════════
// AGGREGATION
// ═══════════════════════════════════════════════════════
const nv = v => parseFloat(v) || 0;
function totals(weeks) {
  const ws = Object.values(weeks || {});
  return {
    membership:   Math.max(0,...ws.map(w=>nv(w.det.membership))),
    attendance:   ws.reduce((a,w)=>a+nv(w.det.attendance),0),
    awareness:    ws.reduce((a,w)=>a+nv(w.wit.awareness),0),
    exposure:     ws.reduce((a,w)=>a+nv(w.wit.exposure),0),
    converts:     ws.reduce((a,w)=>a+nv(w.wit.converts),0),
    numTrained:   ws.reduce((a,w)=>a+nv(w.wit.numTrained),0),
    cdfExisting:  Math.max(0,...ws.map(w=>nv(w.cdf.existing))),
    cdfAdded:     ws.reduce((a,w)=>a+nv(w.cdf.added),0),
    cdfMbr:       Math.max(0,...ws.map(w=>nv(w.cdf.totalMbr))),
    cdfMet:       ws.reduce((a,w)=>a+nv(w.cdf.metWeek),0),
    cdfPresent:   ws.reduce((a,w)=>a+nv(w.cdf.present),0),
    prayerEng:    ws.reduce((a,w)=>a+nv(w.pry.engagements),0),
    inPrayer:     ws.reduce((a,w)=>a+nv(w.pry.inPrayer),0),
    studsTrained: ws.reduce((a,w)=>a+nv(w.pry.studsTrained),0),
    dailyQT:      ws.reduce((a,w)=>a+nv(w.pry.dailyQT),0),
    sgbs:         ws.reduce((a,w)=>a+nv(w.pry.sgbs),0),
    memorization: ws.reduce((a,w)=>a+nv(w.pry.memorization),0),
  };
}
function weekStats(w){
  w = w || {det:{},wit:{},cdf:{},pry:{}};
  return {
    membership:   nv(w.det.membership),
    attendance:   nv(w.det.attendance),
    awareness:    nv(w.wit.awareness),
    exposure:     nv(w.wit.exposure),
    converts:     nv(w.wit.converts),
    numTrained:   nv(w.wit.numTrained),
    cdfExisting:  nv(w.cdf.existing),
    cdfAdded:     nv(w.cdf.added),
    cdfMbr:       nv(w.cdf.totalMbr),
    cdfMet:       nv(w.cdf.metWeek),
    cdfPresent:   nv(w.cdf.present),
    prayerEng:    nv(w.pry.engagements),
    inPrayer:     nv(w.pry.inPrayer),
    studsTrained: nv(w.pry.studsTrained),
    dailyQT:      nv(w.pry.dailyQT),
    sgbs:         nv(w.pry.sgbs),
    memorization: nv(w.pry.memorization),
  };
}
function sumTotals(items){
  const keys=["membership","attendance","awareness","exposure","converts","numTrained","cdfExisting","cdfAdded","cdfMbr","cdfMet","cdfPresent","prayerEng","inPrayer","studsTrained","dailyQT","sgbs","memorization"];
  return items.reduce((acc,t)=>{ if(!t)return acc; keys.forEach(k=>{acc[k]=(acc[k]||0)+(t[k]||0);}); return acc; },{});
}

// ═══════════════════════════════════════════════════════
// GOOGLE SHEETS SYNC
// ═══════════════════════════════════════════════════════
const SHEETS_HEADERS = [
  "Timestamp","Fellowship","Student Name","President","Special Program","Month","Year","Sub. Date",
  "W1 Mbsp","W1 Attd","W1 Prev?","W1 Aware","W1 Exp","W1 Conv","W1 Approach","W1 Train Prog","W1 Trained","W1 Train Topics",
  "W1 CDF Exist","W1 CDF Added","W1 CDF Mbr","W1 CDF Met","W1 CDF Present","W1 Lead Train","W1 CDF Topics","W1 CDF Res",
  "W1 Prayer Eng","W1 In Prayer","W1 Studs Train","W1 Scrip Topics","W1 Daily QT","W1 SGBS","W1 BS Topics","W1 Mem",
  "W2 Mbsp","W2 Attd","W2 Prev?","W2 Aware","W2 Exp","W2 Conv","W2 Approach","W2 Train Prog","W2 Trained","W2 Train Topics",
  "W2 CDF Exist","W2 CDF Added","W2 CDF Mbr","W2 CDF Met","W2 CDF Present","W2 Lead Train","W2 CDF Topics","W2 CDF Res",
  "W2 Prayer Eng","W2 In Prayer","W2 Studs Train","W2 Scrip Topics","W2 Daily QT","W2 SGBS","W2 BS Topics","W2 Mem",
  "W3 Mbsp","W3 Attd","W3 Prev?","W3 Aware","W3 Exp","W3 Conv","W3 Approach","W3 Train Prog","W3 Trained","W3 Train Topics",
  "W3 CDF Exist","W3 CDF Added","W3 CDF Mbr","W3 CDF Met","W3 CDF Present","W3 Lead Train","W3 CDF Topics","W3 CDF Res",
  "W3 Prayer Eng","W3 In Prayer","W3 Studs Train","W3 Scrip Topics","W3 Daily QT","W3 SGBS","W3 BS Topics","W3 Mem",
  "W4 Mbsp","W4 Attd","W4 Prev?","W4 Aware","W4 Exp","W4 Conv","W4 Approach","W4 Train Prog","W4 Trained","W4 Train Topics",
  "W4 CDF Exist","W4 CDF Added","W4 CDF Mbr","W4 CDF Met","W4 CDF Present","W4 Lead Train","W4 CDF Topics","W4 CDF Res",
  "W4 Prayer Eng","W4 In Prayer","W4 Studs Train","W4 Scrip Topics","W4 Daily QT","W4 SGBS","W4 BS Topics","W4 Mem",
  "W5 Mbsp","W5 Attd","W5 Prev?","W5 Aware","W5 Exp","W5 Conv","W5 Approach","W5 Train Prog","W5 Trained","W5 Train Topics",
  "W5 CDF Exist","W5 CDF Added","W5 CDF Mbr","W5 CDF Met","W5 CDF Present","W5 Lead Train","W5 CDF Topics","W5 CDF Res",
  "W5 Prayer Eng","W5 In Prayer","W5 Studs Train","W5 Scrip Topics","W5 Daily QT","W5 SGBS","W5 BS Topics","W5 Mem",
  "Total Mbsp","Total Attd","Total Aware","Total Exp","Total Conv","Total Trained",
  "CDF Existing","CDF Added","CDF Mbr","CDF Met","CDF Present",
  "Prayer Eng","In Prayer","Studs Trained","Daily QT","SGBS","Memorization",
  "FRESH","FINAL","LLTS","Lead Block","Integrity",
  "Staff Visits","Assoc RP","MSPC","Alumni","Quota","Discipled"
];

function buildSheetsRow(report) {
  const h = report.hdr;
  const t = totals(report.weeks);
  const l = report.lead;
  const o = report.other;
  const weekFlat = WEEKS.flatMap(wk => {
    const w = report.weeks[wk];
    return [
      w.det.membership,w.det.attendance,w.det.submittedPrev,
      w.wit.awareness,w.wit.exposure,w.wit.converts,w.wit.approach,w.wit.trainProg,w.wit.numTrained,w.wit.trainTopics,
      w.cdf.existing,w.cdf.added,w.cdf.totalMbr,w.cdf.metWeek,w.cdf.present,w.cdf.leadTrained,w.cdf.trainTopics,w.cdf.resourcePersons,
      w.pry.engagements,w.pry.inPrayer,w.pry.studsTrained,w.pry.trainTopics,w.pry.dailyQT,w.pry.sgbs,w.pry.sgbsTopics,w.pry.memorization
    ];
  });
  return [
    new Date().toISOString(), h.fellowship, h.studentName, h.presidentName, h.specialProg, h.month, h.year, h.subDate,
    ...weekFlat,
    t.membership,t.attendance,t.awareness,t.exposure,t.converts,t.numTrained,
    t.cdfExisting,t.cdfAdded,t.cdfMbr,t.cdfMet,t.cdfPresent,
    t.prayerEng,t.inPrayer,t.studsTrained,t.dailyQT,t.sgbs,t.memorization,
    l.fresh,l.final,l.llts,l.leadBlock,l.integrity,
    o.staffVisits,o.assocRP,o.mspc,o.alumni,o.quota,o.discipled
  ];
}

async function syncToSheets(scriptUrl, report, sheetName) {
  if (!scriptUrl) return { ok: false, msg: "No script URL configured" };
  try {
    const row = buildSheetsRow(report);
    await fetch(scriptUrl, {
      method:"POST", mode:"no-cors",
      headers:{"Content-Type":"application/json"},
      body: JSON.stringify({ sheetName: sheetName || `${report.hdr.month} ${report.hdr.year}`, headers: SHEETS_HEADERS, rows:[row] })
    });
    return { ok: true, msg: "Synced to Google Sheets ✓" };
  } catch(e) {
    return { ok: false, msg: "Sync failed: " + e.message };
  }
}

// ═══════════════════════════════════════════════════════
// THEMES
// ═══════════════════════════════════════════════════════
const DARK = {
  mode:"dark",
  bg:"#080E1A", bg2:"#0D1525", card:"rgba(255,255,255,0.04)", cardSolid:"#0F1C2E",
  border:"rgba(255,255,255,0.09)", borderFocus:"#16A34A",
  headerBg:"#080E1A", headerBorder:"rgba(255,255,255,0.09)",
  navBg:"rgba(8,14,26,0.95)", navActiveBorder:"#16A34A",
  text:"#EFF6FF", muted:"#7C92B1", sub:"#334155",
  green:"#16A34A", greenLt:"#22C55E", glow:"rgba(22,163,74,0.15)",
  gold:"#CA8A04", goldLt:"#FDE68A",
  inputBg:"rgba(255,255,255,0.06)", selectBg:"#0D1525",
  scrollThumb:"rgba(255,255,255,0.12)",
  wit:"#F59E0B",  witBg:"rgba(245,158,11,0.12)",
  disc:"#3B82F6", discBg:"rgba(59,130,246,0.12)",
  pry:"#8B5CF6",  pryBg:"rgba(139,92,246,0.12)",
  lead:"#EC4899", leadBg:"rgba(236,72,153,0.12)",
  prog:"#06B6D4", progBg:"rgba(6,182,212,0.12)",
  oth:"#94A3B8",  othBg:"rgba(148,163,184,0.12)",
  statBg: (c) => c+"1A",
};
const BRIGHT = {
  mode:"bright",
  bg:"#EBF5FF", bg2:"#DBEAFE", card:"#FFFFFF", cardSolid:"#FFFFFF",
  border:"#BFDBFE", borderFocus:"#1D4ED8",
  headerBg:"#1E3A6E", headerBorder:"#1E3A6E",
  navBg:"#FFFFFF", navActiveBorder:"#1D4ED8",
  text:"#0F172A", muted:"#334155", sub:"#64748B",
  green:"#15803D", greenLt:"#16A34A", glow:"rgba(21,128,61,0.12)",
  gold:"#B45309", goldLt:"#92400E",
  inputBg:"#F0F7FF", selectBg:"#EBF5FF",
  scrollThumb:"rgba(30,58,110,0.2)",
  wit:"#DC2626",  witBg:"rgba(220,38,38,0.10)",
  disc:"#1D4ED8", discBg:"rgba(29,78,216,0.10)",
  pry:"#7C3AED",  pryBg:"rgba(124,58,237,0.10)",
  lead:"#DB2777", leadBg:"rgba(219,39,119,0.10)",
  prog:"#0891B2", progBg:"rgba(8,145,178,0.10)",
  oth:"#475569",  othBg:"rgba(71,85,105,0.10)",
  statBg: (c) => c+"18",
};

// ═══════════════════════════════════════════════════════
// UI PRIMITIVES
// ═══════════════════════════════════════════════════════
function Lbl({c,T}){ return <div style={{fontSize:10.5,fontWeight:700,color:T.muted,textTransform:"uppercase",letterSpacing:"0.6px",marginBottom:5}}>{c}</div>; }
function Field({label,children,mb=12,T}){ return <div style={{marginBottom:mb}}>{label&&<Lbl c={label} T={T}/>}{children}</div>; }

function inpStyle(T,extra={}){ return {width:"100%",background:T.inputBg,border:`1px solid ${T.border}`,borderRadius:8,color:T.text,padding:"9px 11px",fontSize:13.5,outline:"none",boxSizing:"border-box",fontFamily:"inherit",...extra}; }

function NIn({label,v,set,mb,T}){
  const handleChange=(e)=>{
    const val=e.target.value;
    if(val===""){set("");return;}
    if(/^\d*\.?\d*$/.test(val))set(val); // digits only — blocks "-" and stray letters
  };
  return <Field label={label} mb={mb} T={T}>
    <input type="number" min="0" inputMode="numeric" value={v} onChange={handleChange} style={inpStyle(T)}
      onFocus={e=>e.target.style.borderColor=T.borderFocus} onBlur={e=>e.target.style.borderColor=T.border}/>
  </Field>;
}
function TIn({label,v,set,ph,multi,mb,T}){
  const s=inpStyle(T,multi?{resize:"vertical",minHeight:58}:{});
  return <Field label={label} mb={mb} T={T}>
    {multi?<textarea value={v} onChange={e=>set(e.target.value)} placeholder={ph} style={s}
        onFocus={e=>e.target.style.borderColor=T.borderFocus} onBlur={e=>e.target.style.borderColor=T.border}/>
      :<input type="text" value={v} onChange={e=>set(e.target.value)} placeholder={ph} style={s}
          onFocus={e=>e.target.style.borderColor=T.borderFocus} onBlur={e=>e.target.style.borderColor=T.border}/>}
  </Field>;
}
function Sel({label,v,set,opts,mb,T}){
  return <Field label={label} mb={mb} T={T}>
    <select value={v} onChange={e=>set(e.target.value)} style={inpStyle(T,{cursor:"pointer",background:T.selectBg})}
      onFocus={e=>e.target.style.borderColor=T.borderFocus} onBlur={e=>e.target.style.borderColor=T.border}>
      {opts.map(o=><option key={o} value={o}>{o}</option>)}
    </select>
  </Field>;
}
function YN({label,v,set,T}){
  return <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",
    padding:"9px 0",borderBottom:`1px solid ${T.border}`}}>
    <span style={{fontSize:13,color:T.text}}>{label}</span>
    <div style={{display:"flex",gap:6}}>
      {["YES","NO"].map(o=>(
        <button key={o} onClick={()=>set(o)} style={{
          padding:"4px 14px",borderRadius:20,fontSize:12,fontWeight:700,cursor:"pointer",border:"none",
          background:v===o?(o==="YES"?T.green:"#DC2626"):(T.mode==="bright"?"#E2E8F0":"rgba(255,255,255,0.08)"),
          color:v===o?"#fff":T.muted,transition:"all .15s"}}>{o}</button>
      ))}
    </div>
  </div>;
}
function Card({children,accent,style:st,T}){
  return <div className="app-card" style={{background:T.card,border:`1px solid ${accent?accent+"55":T.border}`,
    borderRadius:12,padding:"16px 18px",marginBottom:12,
    boxShadow:T.mode==="bright"?"0 1px 6px rgba(0,0,0,0.08)":"none",
    ...(accent&&{borderLeft:`3px solid ${accent}`}),...st}}>{children}</div>;
}
function SecHead({title,color,icon,bg,T}){
  return <div style={{display:"flex",alignItems:"center",gap:10,
    background:bg||color+(T.mode==="bright"?"18":"1A"),borderRadius:"10px 10px 0 0",
    padding:"11px 16px",marginBottom:14,marginTop:6,borderBottom:`2px solid ${color}55`}}>
    <span style={{fontSize:19}}>{icon}</span>
    <span style={{fontSize:12,fontWeight:800,color,textTransform:"uppercase",letterSpacing:"1px"}}>{title}</span>
  </div>;
}
function StatCard({label,value,color,T}){
  return <div style={{background:T.statBg(color),border:`1px solid ${color}33`,
    borderRadius:10,padding:"12px 14px",textAlign:"center",
    boxShadow:T.mode==="bright"?"0 1px 4px rgba(0,0,0,0.06)":"none"}}>
    <div style={{fontSize:26,fontWeight:800,color,lineHeight:1}}>{nv(value)||"—"}</div>
    <div style={{fontSize:10,color:T.muted,marginTop:4,textTransform:"uppercase",letterSpacing:"0.5px"}}>{label}</div>
  </div>;
}
function Grid({cols,gap=10,children,style:st}){ return <div className="app-grid" style={{display:"grid",gridTemplateColumns:cols,gap,...st}}>{children}</div>; }

// ═══════════════════════════════════════════════════════
// APPS SCRIPT CODE (shown in settings)
// ═══════════════════════════════════════════════════════
const APPS_SCRIPT = `// ── GHAFES Ministry Report Sync ──────────────────────
// 1. Open script.google.com and create a new project
// 2. Paste this entire script
// 3. Replace YOUR_SPREADSHEET_ID below with your Sheet ID
//    (found in the Sheet URL: .../d/SHEET_ID/edit)
// 4. Click Deploy → New deployment → Web app
// 5. Set: Execute as = Me, Access = Anyone
// 6. Copy the deployment URL and paste it in Settings

const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE";

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(SPREADSHEET_ID);
    const name = data.sheetName || "Ministry Reports";
    let sheet  = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      sheet.appendRow(data.headers || []);
      sheet.getRange(1,1,1,sheet.getLastColumn())
           .setFontWeight("bold").setBackground("#1E3A6E")
           .setFontColor("#FFFFFF");
      sheet.setFrozenRows(1);
    }
    (data.rows || []).forEach(r => sheet.appendRow(r));
    const out = { success:true, sheet:name, rows:(data.rows||[]).length };
    return ContentService
      .createTextOutput(JSON.stringify(out))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({success:false,error:err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
function doGet(){ return ContentService.createTextOutput("GHAFES Sync Active"); }`;

// ═══════════════════════════════════════════════════════
// SETTINGS MODAL
// ═══════════════════════════════════════════════════════
// ═══════════════════════════════════════════════════════
// STAFF ACCESS GATE — password-unlocks CMC / Zonal / Sheets Settings
// ═══════════════════════════════════════════════════════
function StaffGateModal({unlocked,onUnlock,onClose,onGoTo,T}){
  const ROLES = [
    {id:"cmc",    label:"CMC Report",           desc:"For Campus Ministry Coordinators"},
    {id:"zonal",  label:"Zonal Report",         desc:"For Zonal Coordinators"},
    {id:"sheets", label:"Google Sheets Settings", desc:"For the person managing the data sync"},
  ];
  const [pw,setPw]     = useState({cmc:"",zonal:"",sheets:""});
  const [err,setErr]   = useState({});
  const [busy,setBusy] = useState({});

  const boxSt = { position:"fixed",inset:0,background:"rgba(0,0,0,0.65)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16 };
  const panelSt = { background:T.mode==="bright"?"#FFFFFF":"#0D1525",border:`1px solid ${T.border}`,borderRadius:16,width:"100%",maxWidth:440,boxShadow:"0 24px 80px rgba(0,0,0,0.5)" };

  const attempt=async(role)=>{
    setBusy(b=>({...b,[role]:true})); setErr(e=>({...e,[role]:""}));
    const ok = await checkAccess(role,pw[role]);
    setBusy(b=>({...b,[role]:false}));
    if(ok){ onUnlock(role); onGoTo(role); }
    else setErr(e=>({...e,[role]:"Incorrect password."}));
  };

  return <div style={boxSt} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
    <div style={panelSt}>
      <div style={{background:"#1E3A6E",padding:"18px 24px",borderRadius:"16px 16px 0 0",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div>
          <div style={{fontSize:16,fontWeight:800,color:"#fff"}}>🔒 Staff Access</div>
          <div style={{fontSize:12,color:"rgba(255,255,255,0.65)",marginTop:3}}>Enter the password for the area you need</div>
        </div>
        <button onClick={onClose} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:18,fontWeight:700,fontFamily:"inherit"}}>×</button>
      </div>
      <div style={{padding:22,display:"flex",flexDirection:"column",gap:16}}>
        {ROLES.map(r=>{
          const isUnlocked = unlocked[r.id];
          return <div key={r.id} style={{border:`1px solid ${T.border}`,borderRadius:10,padding:"12px 14px"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:isUnlocked?0:8}}>
              <div>
                <div style={{fontSize:13,fontWeight:700,color:T.text}}>{r.label}</div>
                <div style={{fontSize:11,color:T.muted}}>{r.desc}</div>
              </div>
              {isUnlocked&&<span style={{fontSize:11,fontWeight:700,color:T.green,background:T.glow,padding:"3px 10px",borderRadius:20}}>✓ Unlocked</span>}
            </div>
            {!isUnlocked&&<>
              <div style={{display:"flex",gap:8}}>
                <input type="password" value={pw[r.id]} onChange={e=>setPw(p=>({...p,[r.id]:e.target.value}))}
                  onKeyDown={e=>{if(e.key==="Enter")attempt(r.id);}}
                  placeholder="Password" style={inpStyle(T,{flex:1})}/>
                <button onClick={()=>attempt(r.id)} disabled={busy[r.id]} style={{
                  background:T.mode==="bright"?"#1E3A6E":T.glow,border:`1.5px solid ${T.mode==="bright"?"#1E3A6E":T.green}`,
                  color:T.mode==="bright"?"#fff":T.green,padding:"0 16px",borderRadius:8,fontSize:12.5,fontWeight:700,
                  cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>
                  {busy[r.id]?"…":"Unlock"}
                </button>
              </div>
              {err[r.id]&&<div style={{fontSize:11.5,color:"#DC2626",marginTop:6}}>{err[r.id]}</div>}
            </>}
            {isUnlocked&&<button onClick={()=>onGoTo(r.id)} style={{marginTop:8,background:"transparent",border:`1px solid ${T.border}`,color:T.text,padding:"7px 14px",borderRadius:8,fontSize:12,fontWeight:600,cursor:"pointer",fontFamily:"inherit"}}>Go to {r.label} →</button>}
          </div>;
        })}
        <div style={{fontSize:10.5,color:T.sub,textAlign:"center"}}>Unlocked areas stay unlocked until you close this browser tab.</div>
      </div>
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════
// PAST REPORTS BROWSER — lets a fellowship see which months
// already have a saved report, and jump straight to one
// ═══════════════════════════════════════════════════════
function PastReportsModal({fellowship,allReports,onPick,onClose,T}){
  const boxSt = { position:"fixed",inset:0,background:"rgba(0,0,0,0.65)",zIndex:1000,display:"flex",alignItems:"center",justifyContent:"center",padding:16 };
  const panelSt = { background:T.mode==="bright"?"#FFFFFF":"#0D1525",border:`1px solid ${T.border}`,borderRadius:16,width:"100%",maxWidth:420,maxHeight:"80vh",overflow:"auto",boxShadow:"0 24px 80px rgba(0,0,0,0.5)" };
  const rows=[];
  [...YEARS].reverse().forEach(year=>MONTHS.forEach(month=>{
    const r=allReports[rKey(fellowship,month,year)];
    if(r)rows.push({month,year,membership:totals(r.weeks).membership||0});
  }));
  return <div style={boxSt} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
    <div style={panelSt}>
      <div style={{background:"#1E3A6E",padding:"18px 24px",borderRadius:"16px 16px 0 0",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div>
          <div style={{fontSize:16,fontWeight:800,color:"#fff"}}>📚 Past Reports</div>
          <div style={{fontSize:12,color:"rgba(255,255,255,0.65)",marginTop:3}}>{fellowship}</div>
        </div>
        <button onClick={onClose} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:18,fontWeight:700,fontFamily:"inherit"}}>×</button>
      </div>
      <div style={{padding:18}}>
        {rows.length===0
          ?<div style={{fontSize:12.5,color:T.muted,textAlign:"center",padding:"20px 8px"}}>No saved reports yet for this fellowship.</div>
          :<div style={{display:"flex",flexDirection:"column",gap:6}}>
            {rows.map(r=>(
              <button key={r.month+r.year} onClick={()=>{onPick(r.month,r.year);onClose();}} style={{
                display:"flex",alignItems:"center",justifyContent:"space-between",width:"100%",textAlign:"left",
                background:T.mode==="bright"?"#F8FAFC":"rgba(255,255,255,0.04)",border:`1px solid ${T.border}`,
                borderRadius:9,padding:"10px 14px",cursor:"pointer",fontFamily:"inherit"}}>
                <span style={{fontSize:13,fontWeight:700,color:T.text}}>{r.month} {r.year}</span>
                <span style={{fontSize:11,color:T.muted}}>{r.membership} members reported →</span>
              </button>
            ))}
          </div>}
      </div>
    </div>
  </div>;
}

function SettingsModal({settings,onSave,onClose,T}){
  const [url,setUrl]   = useState(settings.scriptUrl||"");
  const [name,setName] = useState(settings.sheetName||"Ministry Reports");
  const [copied,setCopied] = useState(false);

  const handleCopy = ()=>{
    navigator.clipboard.writeText(APPS_SCRIPT).then(()=>{setCopied(true);setTimeout(()=>setCopied(false),2000);});
  };

  const boxSt = {
    position:"fixed",inset:0,background:"rgba(0,0,0,0.65)",zIndex:1000,
    display:"flex",alignItems:"center",justifyContent:"center",padding:16
  };
  const panelSt = {
    background:T.mode==="bright"?"#FFFFFF":"#0D1525",
    border:`1px solid ${T.border}`,borderRadius:16,
    width:"100%",maxWidth:680,maxHeight:"90vh",overflow:"auto",
    boxShadow:"0 24px 80px rgba(0,0,0,0.5)"
  };

  return <div style={boxSt} onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
    <div style={panelSt}>
      {/* Modal header */}
      <div style={{background:"#1E3A6E",padding:"18px 24px",borderRadius:"16px 16px 0 0",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <div>
          <div style={{fontSize:16,fontWeight:800,color:"#fff"}}>⚙ Settings — Google Sheets Integration</div>
          <div style={{fontSize:12,color:"rgba(255,255,255,0.65)",marginTop:3}}>Sync reports automatically to your Google Spreadsheet</div>
        </div>
        <button onClick={onClose} style={{background:"rgba(255,255,255,0.15)",border:"none",color:"#fff",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:18,fontWeight:700,fontFamily:"inherit"}}>×</button>
      </div>

      <div style={{padding:24}}>
        {/* Step 1 */}
        <div style={{marginBottom:20}}>
          <div style={{fontSize:13,fontWeight:800,color:T.mode==="bright"?"#1E3A6E":T.disc,marginBottom:10,display:"flex",alignItems:"center",gap:8}}>
            <span style={{background:T.mode==="bright"?"#1E3A6E":T.disc,color:"#fff",width:22,height:22,borderRadius:"50%",display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:12,flexShrink:0}}>1</span>
            Set Up Your Google Apps Script
          </div>
          <div style={{fontSize:12,color:T.muted,marginBottom:10,lineHeight:1.6}}>
            Copy the script below → go to <strong style={{color:T.mode==="bright"?"#1E3A6E":T.disc}}>script.google.com</strong> → create new project → paste &amp; replace <code style={{background:T.mode==="bright"?"#EBF5FF":"rgba(255,255,255,0.08)",padding:"1px 5px",borderRadius:4,fontSize:11}}>YOUR_SPREADSHEET_ID_HERE</code> → Deploy as Web App.
          </div>
          <div style={{position:"relative"}}>
            <pre style={{
              background:T.mode==="bright"?"#0F172A":"#020810",color:"#A3E635",
              fontSize:11,padding:"14px 16px",borderRadius:10,overflow:"auto",
              maxHeight:200,border:`1px solid ${T.border}`,margin:0,lineHeight:1.5,
              fontFamily:"'Courier New',monospace"
            }}>{APPS_SCRIPT}</pre>
            <button onClick={handleCopy} style={{
              position:"absolute",top:8,right:8,
              background:copied?"#15803D":"rgba(163,230,53,0.15)",
              border:`1px solid ${copied?"#15803D":"#A3E635"}`,
              color:copied?"#fff":"#A3E635",
              padding:"5px 12px",borderRadius:6,fontSize:11,fontWeight:700,cursor:"pointer",fontFamily:"inherit"
            }}>{copied?"✓ Copied!":"Copy Script"}</button>
          </div>
        </div>

        {/* Step 2 */}
        <div style={{marginBottom:20}}>
          <div style={{fontSize:13,fontWeight:800,color:T.mode==="bright"?"#1E3A6E":T.disc,marginBottom:10,display:"flex",alignItems:"center",gap:8}}>
            <span style={{background:T.mode==="bright"?"#1E3A6E":T.disc,color:"#fff",width:22,height:22,borderRadius:"50%",display:"inline-flex",alignItems:"center",justifyContent:"center",fontSize:12,flexShrink:0}}>2</span>
            Configure Connection
          </div>
          <div style={{display:"grid",gridTemplateColumns:"2fr 1fr",gap:12}}>
            <Field label="Apps Script Deployment URL" T={T}>
              <input type="url" value={url} onChange={e=>setUrl(e.target.value)}
                placeholder="https://script.google.com/macros/s/AKf.../exec"
                style={inpStyle(T)}
                onFocus={e=>e.target.style.borderColor=T.borderFocus} onBlur={e=>e.target.style.borderColor=T.border}/>
            </Field>
            <Field label="Sheet Name Prefix" T={T}>
              <input type="text" value={name} onChange={e=>setName(e.target.value)}
                placeholder="Ministry Reports"
                style={inpStyle(T)}
                onFocus={e=>e.target.style.borderColor=T.borderFocus} onBlur={e=>e.target.style.borderColor=T.border}/>
            </Field>
          </div>
          <div style={{fontSize:11,color:T.muted,marginTop:4}}>The sheet name will be: <em>{name || "Ministry Reports"} — Month Year</em> (e.g., "{name || "Ministry Reports"} — July 2026")</div>
        </div>

        {/* Info box */}
        <div style={{background:T.mode==="bright"?"#EFF7FF":"rgba(29,78,216,0.1)",border:`1px solid ${T.disc}44`,borderRadius:10,padding:"12px 16px",marginBottom:20,fontSize:12,color:T.text,lineHeight:1.6}}>
          📌 <strong>How sync works:</strong> Each time you click "Save &amp; Sync Report", the fellowship report is saved locally <em>and</em> a row is appended to your Google Sheet. No data is ever overwritten — each save appends a new row with a timestamp.
        </div>

        <div style={{display:"flex",justifyContent:"flex-end",gap:10}}>
          <button onClick={onClose} style={{
            background:"transparent",border:`1.5px solid ${T.border}`,color:T.muted,
            padding:"10px 22px",borderRadius:9,fontSize:13,fontWeight:600,cursor:"pointer",fontFamily:"inherit"
          }}>Cancel</button>
          <button onClick={()=>onSave({scriptUrl:url.trim(),sheetName:name.trim()})} style={{
            background:"#1E3A6E",border:"none",color:"#fff",
            padding:"10px 24px",borderRadius:9,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"
          }}>Save Settings</button>
        </div>
      </div>
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════
// FORM SECTIONS
// ═══════════════════════════════════════════════════════
function DetSection({d,set,T}){
  const u=(f,v)=>set({...d,[f]:v});
  return <>
    <SecHead title="Fellowship Details" color={T.oth} icon="👥" T={T}/>
    <Grid cols="1fr 1fr"><NIn label="Membership" v={d.membership} set={v=>u("membership",v)} T={T}/><NIn label="Week's Attendance" v={d.attendance} set={v=>u("attendance",v)} T={T}/></Grid>
    <div style={{borderRadius:8,overflow:"hidden",border:`1px solid ${T.border}`,marginBottom:10}}>
      <YN label="Submitted Previous Reports?" v={d.submittedPrev} set={v=>u("submittedPrev",v)} T={T}/>
    </div>
    {d.submittedPrev==="NO"&&<Grid cols="1fr 1fr">
      <TIn label="Why Not?" v={d.whyNot} set={v=>u("whyNot",v)} ph="Reason for non-submission" T={T}/>
      <TIn label="When to Expect Report?" v={d.whenExpect} set={v=>u("whenExpect",v)} ph="Expected date" T={T}/>
    </Grid>}
  </>;
}
function WitSection({d,set,T}){
  const u=(f,v)=>set({...d,[f]:v});
  return <>
    <SecHead title="Thriving in Witness" color={T.wit} icon="🔥" bg={T.witBg} T={T}/>
    <Grid cols="1fr 1fr 1fr">
      <NIn label="Evangelism Awareness" v={d.awareness} set={v=>u("awareness",v)} T={T}/>
      <NIn label="Gospel Exposure" v={d.exposure} set={v=>u("exposure",v)} T={T}/>
      <NIn label="New Converts & Rededicants" v={d.converts} set={v=>u("converts",v)} T={T}/>
    </Grid>
    <TIn label="Evangelism / Mission Approach(es)" v={d.approach} set={v=>u("approach",v)} ph="Describe approaches used this week" multi T={T}/>
    <Grid cols="1fr 1fr 1fr">
      <TIn label="Evangelism Training / Equipping Program" v={d.trainProg} set={v=>u("trainProg",v)} ph="None if N/A" T={T}/>
      <NIn label="Number Trained" v={d.numTrained} set={v=>u("numTrained",v)} T={T}/>
      <TIn label="Evangelism Training Topics" v={d.trainTopics} set={v=>u("trainTopics",v)} ph="Topics covered" T={T}/>
    </Grid>
  </>;
}
function CdfSection({d,set,T}){
  const u=(f,v)=>set({...d,[f]:v});
  return <>
    <SecHead title="Christlike Discipleship Families (CDF)" color={T.disc} icon="👨‍👩‍👧‍👦" bg={T.discBg} T={T}/>
    <Grid cols="1fr 1fr 1fr">
      <NIn label="No. of Existing CDFs" v={d.existing} set={v=>u("existing",v)} T={T}/>
      <NIn label="New CDFs Added" v={d.added} set={v=>u("added",v)} T={T}/>
      <NIn label="Total CDF Membership" v={d.totalMbr} set={v=>u("totalMbr",v)} T={T}/>
    </Grid>
    <Grid cols="1fr 1fr">
      <NIn label="No. of CDFs that Met this Week" v={d.metWeek} set={v=>u("metWeek",v)} T={T}/>
      <NIn label="CDF Members Present this Week" v={d.present} set={v=>u("present",v)} T={T}/>
    </Grid>
    <div style={{borderRadius:8,overflow:"hidden",border:`1px solid ${T.border}`,marginBottom:10}}>
      <YN label="CDF Leaders Trained?" v={d.leadTrained} set={v=>u("leadTrained",v)} T={T}/>
    </div>
    {d.leadTrained==="YES"&&<Grid cols="1fr 1fr">
      <TIn label="CDF Training Topics" v={d.trainTopics} set={v=>u("trainTopics",v)} ph="Topics covered" T={T}/>
      <TIn label="Resource Persons" v={d.resourcePersons} set={v=>u("resourcePersons",v)} ph="Name(s)" T={T}/>
    </Grid>}
  </>;
}
function PrySection({d,set,T}){
  const u=(f,v)=>set({...d,[f]:v});
  return <>
    <SecHead title="Scripture & Prayer Engagement" color={T.pry} icon="📖" bg={T.pryBg} T={T}/>
    <Grid cols="1fr 1fr">
      <NIn label="No. of Prayer Engagements" v={d.engagements} set={v=>u("engagements",v)} T={T}/>
      <NIn label="No. Involved in Prayer" v={d.inPrayer} set={v=>u("inPrayer",v)} T={T}/>
    </Grid>
    <Grid cols="1fr 1fr">
      <NIn label="Students Trained (Scripture Engagement)" v={d.studsTrained} set={v=>u("studsTrained",v)} T={T}/>
      <TIn label="Scripture Eng. Training Topics" v={d.trainTopics} set={v=>u("trainTopics",v)} ph="Topics covered" T={T}/>
    </Grid>
    <Grid cols="1fr 1fr 1fr">
      <NIn label="Students Having Daily Quiet Time" v={d.dailyQT} set={v=>u("dailyQT",v)} T={T}/>
      <NIn label="Small Group Bible Study" v={d.sgbs} set={v=>u("sgbs",v)} T={T}/>
      <NIn label="Scripture Memorization" v={d.memorization} set={v=>u("memorization",v)} T={T}/>
    </Grid>
    <TIn label="Small Group Bible Study Topics" v={d.sgbsTopics} set={v=>u("sgbsTopics",v)} ph="Topics studied" mb={4} T={T}/>
  </>;
}
function LeadSection({d,set,T}){
  const u=(f,v)=>set({...d,[f]:v});
  const items=[["fresh","FRESH"],["final","FINAL"],["llts","LLTS"],["leadBlock","LEADERSHIP BLOCK"],["integrity","INTEGRITY SERIES"]];
  return <>
    <SecHead title="Leadership Development & Mentorship" color={T.lead} icon="🏆" bg={T.leadBg} T={T}/>
    <div style={{borderRadius:8,overflow:"hidden",border:`1px solid ${T.border}`,marginBottom:8}}>
      {items.map(([f,l])=><YN key={f} label={l} v={d[f]} set={v=>u(f,v)} T={T}/>)}
    </div>
    <div style={{fontSize:11,color:T.muted,fontStyle:"italic",marginBottom:4}}>* Provide details of YES items in the Programs section</div>
  </>;
}
function ProgSection({zonalProg,majorProg,programs,setZonal,setMajor,setProgs,T}){
  const addProg=()=>setProgs([...programs,{name:"",level:"Local",hostLCF:"",participants:"",exposure:"",converts:"",resourcePersons:"",testimonies:[""]}]);
  const remProg=(i)=>setProgs(programs.filter((_,j)=>j!==i));
  const updProg=(i,f,v)=>setProgs(programs.map((p,j)=>j===i?{...p,[f]:v}:p));
  const addTestimony=(i)=>setProgs(programs.map((p,j)=>j===i?{...p,testimonies:[...(p.testimonies||[]),""]}:p));
  const updTestimony=(i,k,v)=>setProgs(programs.map((p,j)=>j===i?{...p,testimonies:(p.testimonies||[]).map((t,l)=>l===k?v:t)}:p));
  const remTestimony=(i,k)=>setProgs(programs.map((p,j)=>j===i?{...p,testimonies:(p.testimonies||[]).filter((_,l)=>l!==k)}:p));
  return <>
    <SecHead title="Zonal, National & Major Programs" color={T.prog} icon="🌐" bg={T.progBg} T={T}/>
    <div style={{borderRadius:8,overflow:"hidden",border:`1px solid ${T.border}`,marginBottom:14}}>
      <YN label="Was there any Zonal or National Program this week?" v={zonalProg} set={setZonal} T={T}/>
      <YN label="Was there any Major Program held by your fellowship?" v={majorProg} set={setMajor} T={T}/>
    </div>
    {programs.map((p,i)=>{
      const testimonies=p.testimonies&&p.testimonies.length?p.testimonies:[""];
      return <Card key={i} accent={T.prog} T={T}>
        <div style={{display:"flex",justifyContent:"space-between",marginBottom:10}}>
          <span style={{fontSize:12,fontWeight:700,color:T.prog,textTransform:"uppercase",letterSpacing:"0.6px"}}>Activity {i+1}</span>
          <button onClick={()=>remProg(i)} style={{background:"none",border:"none",color:"#DC2626",cursor:"pointer",fontSize:18,lineHeight:1}}>×</button>
        </div>
        <Grid cols="2fr 1fr">
          <TIn label="Name of Activity" v={p.name} set={v=>updProg(i,"name",v)} ph="Activity name" T={T}/>
          <Sel label="Level" v={p.level} set={v=>updProg(i,"level",v)} opts={LEVELS} T={T}/>
        </Grid>
        <Grid cols="1fr 1fr 1fr">
          <TIn label="Organizing / Host LCF" v={p.hostLCF} set={v=>updProg(i,"hostLCF",v)} ph="Fellowship name" T={T}/>
          <NIn label="No. of Participants" v={p.participants} set={v=>updProg(i,"participants",v)} T={T}/>
          <NIn label="No. Exposed to Gospel" v={p.exposure} set={v=>updProg(i,"exposure",v)} T={T}/>
        </Grid>
        <Grid cols="1fr 1fr">
          <NIn label="New Converts & Rededicants" v={p.converts} set={v=>updProg(i,"converts",v)} T={T}/>
          <TIn label="Resource Persons (& Topics)" v={p.resourcePersons} set={v=>updProg(i,"resourcePersons",v)} ph="Name (Topic)" T={T}/>
        </Grid>

        <div style={{height:4}}/>
        <Lbl c="Testimonies & Impact Stories" T={T}/>
        {testimonies.map((t,k)=>(
          <div key={k} style={{display:"flex",gap:8,alignItems:"flex-start",marginBottom:8}}>
            <div style={{flex:1}}>
              <TIn v={t} set={v=>updTestimony(i,k,v)} ph={`Testimony / impact story ${k+1}…`} multi T={T} mb={0}/>
            </div>
            {testimonies.length>1&&<button onClick={()=>remTestimony(i,k)} title="Remove this testimony" style={{background:"none",border:"none",color:"#DC2626",cursor:"pointer",fontSize:18,lineHeight:1,marginTop:6,flexShrink:0}}>×</button>}
          </div>
        ))}
        <button onClick={()=>addTestimony(i)} style={{display:"flex",alignItems:"center",gap:6,background:"transparent",border:`1.5px dashed ${T.prog}66`,color:T.prog,borderRadius:8,padding:"7px 14px",fontSize:12,cursor:"pointer",fontFamily:"inherit"}}>＋ Add Another Testimony</button>
      </Card>;
    })}
    {programs.length<4&&<button onClick={addProg} style={{display:"flex",alignItems:"center",gap:6,background:T.progBg,border:`1.5px dashed ${T.prog}66`,color:T.prog,borderRadius:8,padding:"9px 16px",fontSize:12.5,cursor:"pointer",fontFamily:"inherit"}}>＋ Add Activity (max 4)</button>}
  </>;
}
function OthSection({d,set,isCMC,T}){
  const u=(f,v)=>set({...d,[f]:v});
  return <>
    <SecHead title={isCMC?"Other CMC Reports":"Other LCF Reports"} color={T.oth} icon="📋" bg={T.othBg} T={T}/>
    {isCMC
      ?<Grid cols="1fr 1fr"><NIn label="Fellowship Visits" v={d.fellowVisits} set={v=>u("fellowVisits",v)} T={T}/><NIn label="Associate Visits" v={d.assocVisits} set={v=>u("assocVisits",v)} T={T}/><NIn label="Associate Calls" v={d.assocCalls} set={v=>u("assocCalls",v)} T={T}/><NIn label="Associates as Resource Persons" v={d.assocRP} set={v=>u("assocRP",v)} T={T}/><NIn label="Associates Mapped Out" v={d.assocMapped} set={v=>u("assocMapped",v)} T={T}/><NIn label="Students Discipled by Staff" v={d.discipled} set={v=>u("discipled",v)} T={T}/></Grid>
      :<Grid cols="1fr 1fr"><NIn label="Fellowship Visits by Staff" v={d.staffVisits} set={v=>u("staffVisits",v)} T={T}/><NIn label="Associates as Resource Persons" v={d.assocRP} set={v=>u("assocRP",v)} T={T}/><NIn label="MSPC Received" v={d.mspc} set={v=>u("mspc",v)} T={T}/><NIn label="Visits from Fellowship Alumni" v={d.alumni} set={v=>u("alumni",v)} T={T}/><NIn label="Quota for the Month" v={d.quota} set={v=>u("quota",v)} T={T}/><NIn label="Students Discipled by Executives" v={d.discipled} set={v=>u("discipled",v)} T={T}/></Grid>}
  </>;
}

// ═══════════════════════════════════════════════════════
// SECTION NAV
// ═══════════════════════════════════════════════════════
const SECS = (T) => [
  {id:"det",  label:"Fellowship Details",    color:T.oth,  icon:"👥"},
  {id:"wit",  label:"Witness",               color:T.wit,  icon:"🔥"},
  {id:"cdf",  label:"Discipleship (CDFs)",   color:T.disc, icon:"👨‍👩‍👧‍👦"},
  {id:"pry",  label:"Prayer & Scripture",    color:T.pry,  icon:"📖"},
  {id:"lead", label:"Leadership",            color:T.lead, icon:"🏆"},
  {id:"prog", label:"Programs",              color:T.prog, icon:"🌐"},
  {id:"oth",  label:"Other Reports",         color:T.oth,  icon:"📋"},
  {id:"wsum", label:"Weekly Summary",        color:T.green,icon:"📊"},
  {id:"msum", label:"Monthly Summary",       color:T.green,icon:"🗓"},
];

// ═══════════════════════════════════════════════════════
// EXCEL BUILDERS
// ═══════════════════════════════════════════════════════
function buildFellowshipSheet(report){
  const h=report.hdr,l=report.lead,o=report.other;
  const rows=[
    ["MONTHLY REPORT TEMPLATE FOR LOCAL CAMPUS FELLOWSHIPS"],
    ["Student Name:",h.studentName,"","","","","Fellowship:",h.fellowship],
    ["Reporting Period:","","Month:",h.month,"Year:",h.year,"Submission Date:",h.subDate],
    ["Name of Fellowship President:",h.presidentName],
    ["Special Fellowship Program this Week:",h.specialProg],[],
    ["FELLOWSHIP DETAILS"],
    ["Reporting Week","","Membership","Week's Attendance","","Submitted Previous Reports?","Why if not?","","When to expect Report"],
    ...WEEKS.map(wk=>{const d=report.weeks[wk].det;return[wk,"",d.membership,d.attendance,"",d.submittedPrev,d.whyNot,"",d.whenExpect];}),
    [],["THRIVING IN WITNESS"],
    ["Reporting Week","Evangelism Awareness","Gospel Exposure","New Converts & Rededicants","Evangelism/Mission Approach(es)","","Evangelism Training/Equipping Program","","Number Trained","Evangelism Training Topics"],
    ...WEEKS.map(wk=>{const w=report.weeks[wk].wit;return[wk,w.awareness,w.exposure,w.converts,w.approach,"",w.trainProg,"",w.numTrained,w.trainTopics];}),
    [],["THRIVING IN WHOLE LIFE DISCIPLESHIP"],["Christlike Discipleship Families (CDFs)"],
    ["Reporting Week","No. of Existing CDFs","New CDFs Added","Total CDF Membership","No. of CDF that met this Week","CDF Members Present","CDF Leaders Trained?","CDF Training Topics","Resource Persons"],
    ...WEEKS.map(wk=>{const c=report.weeks[wk].cdf;return[wk,c.existing,c.added,c.totalMbr,c.metWeek,c.present,c.leadTrained,c.trainTopics,c.resourcePersons];}),
    [],["Scripture and Prayer Engagement"],
    ["Reporting Week","No. of Prayer Engagements","No. Involved in Prayer","Studs. Trained (Scripture Eng.)","Scripture Eng. Training Topics","Students Having Daily Quiet Time","Small Group Bible Study","Small Group Bible Study Topics","","Scripture Memorization"],
    ...WEEKS.map(wk=>{const p=report.weeks[wk].pry;return[wk,p.engagements,p.inPrayer,p.studsTrained,p.trainTopics,p.dailyQT,p.sgbs,p.sgbsTopics,"",p.memorization];}),
    [],["THRIVING IN LEADERSHIP DEVELOPMENT AND MENTORSHIP"],
    ["FRESH:",l.fresh,"FINAL:",l.final,"LLTS:",l.llts,"LEADERSHIP BLOCK:",l.leadBlock,"INTEGRITY SERIES:",l.integrity],
    [],["ZONAL, NATIONAL AND MAJOR FELLOWSHIP PROGRAMS"],
    ["Was there any Zonal or National Program This week?","","",report.zonalProg,"","Was there any Major Program?","","",report.majorProg],
    [],["Name of Activity","Level","Organizing or Host LCF","Number of Participants","Number Exposed to the Gospel","New Converts and Rededicants","Resource Persons (And Topics)","Testimonies / Impact Stories"],
    ...(report.programs||[]).map(p=>[p.name,p.level,p.hostLCF,p.participants,p.exposure,p.converts,p.resourcePersons,(p.testimonies||[]).filter(t=>t&&t.trim()).join("\n")]),
    [],["TESTIMONIES & IMPACT STORIES (Full List)"],
    ["Activity","#","Testimony / Impact Story"],
    ...(report.programs||[]).flatMap(p=>(p.testimonies||[]).filter(t=>t&&t.trim()).map((t,k)=>[p.name||"(Unnamed Activity)",k+1,t])),
    [],["OTHER LCF REPORTS"],
    ["Fellowship Visits by Staff:",o.staffVisits,"","Associates as Resource Persons:",o.assocRP],
    ["MSPC Received:",o.mspc,"","Visits from Fellowship Alumni:",o.alumni],
    ["Quota for the Month:",o.quota,"","Students Discipled by Executives:",o.discipled],
    [],["THANK YOU FOR YOUR FAITHFULNESS"]
  ];
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"]=Array(10).fill({wch:22}).map((c,i)=>i===7?{wch:50}:c);
  return ws;
}
function buildCMCSheet(cmcReport,fdList){
  const h=cmcReport.hdr,l=cmcReport.lead,o=cmcReport.other;
  const tot=sumTotals(fdList.filter(f=>f.t).map(f=>f.t));
  const rows=[
    ["WEEKLY REPORT TEMPLATE FOR CAMPUS MINISTRY COORDINATORS"],
    ["Staff Name:",h.staffName,"","","","","Station:",h.station],
    ["Reporting Period:","","Month:",h.month,"Year:",h.year,"Submission Date:",h.subDate],
    ["Books Reading this week:","","",h.books],["Devotional Life:","","",h.devLife],[],
    ["FELLOWSHIP DETAILS"],
    ["LCFs Coordinated","","Membership","Month Attendance","","Submitted Report?"],
    ...fdList.map(f=>[f.fellowship,"",f.t?f.t.membership:"",f.t?f.t.attendance:"","",f.t?"YES":"NO"]),
    ["MONTHLY TOTAL","",tot.membership||0,tot.attendance||0],[],
    ["THRIVING IN WITNESS"],
    ["LCF Name","Evangelism Awareness","Gospel Exposure","New Converts & Rededicants","","","","","No. Trained"],
    ...fdList.map(f=>[f.fellowship,f.t?f.t.awareness:0,f.t?f.t.exposure:0,f.t?f.t.converts:0,"","","","",f.t?f.t.numTrained:0]),
    ["MONTHLY TOTAL",tot.awareness||0,tot.exposure||0,tot.converts||0,"","","","",tot.numTrained||0],[],
    ["THRIVING IN WHOLE LIFE DISCIPLESHIP"],["Christlike Discipleship Families (CDFs)"],
    ["LCF Name","No. of Existing CDFs","New CDFs Added","Total CDF Membership","CDF Meetings Held","CDF Members Present"],
    ...fdList.map(f=>[f.fellowship,f.t?f.t.cdfExisting:0,f.t?f.t.cdfAdded:0,f.t?f.t.cdfMbr:0,f.t?f.t.cdfMet:0,f.t?f.t.cdfPresent:0]),
    ["MONTHLY TOTAL",tot.cdfExisting||0,tot.cdfAdded||0,tot.cdfMbr||0,tot.cdfMet||0,tot.cdfPresent||0],[],
    ["Scripture and Prayer Engagement"],
    ["LCF Name","Prayer Engagements","Involved in Prayer","Scripture Training","Daily QT","Small Group Bible Study","","Scripture Memorization"],
    ...fdList.map(f=>[f.fellowship,f.t?f.t.prayerEng:0,f.t?f.t.inPrayer:0,f.t?f.t.studsTrained:0,f.t?f.t.dailyQT:0,f.t?f.t.sgbs:0,"",f.t?f.t.memorization:0]),
    ["MONTHLY TOTAL",tot.prayerEng||0,tot.inPrayer||0,tot.studsTrained||0,tot.dailyQT||0,tot.sgbs||0,"",tot.memorization||0],[],
    ["THRIVING IN LEADERSHIP DEVELOPMENT AND MENTORSHIP"],
    ["FRESH:",l.fresh,"FINAL:",l.final,"LLTS:",l.llts,"LEADERSHIP BLOCK:",l.leadBlock,"INTEGRITY SERIES:",l.integrity],[],
    ["ZONAL, NATIONAL AND MAJOR FELLOWSHIP PROGRAMS"],
    ["Was there any Zonal or National Program?","","",cmcReport.zonalProg,"","Was there any Major Program?","","",cmcReport.majorProg],[],
    ["Name of Activity","Level","Organizing or Host LCF","Participants","Gospel Exposure","New Converts","Resource Persons","Testimonies / Impact Stories"],
    ...(cmcReport.programs||[]).map(p=>[p.name,p.level,p.hostLCF,p.participants,p.exposure,p.converts,p.resourcePersons,(p.testimonies||[]).filter(t=>t&&t.trim()).join("\n")]),
    [],["TESTIMONIES & IMPACT STORIES (Full List)"],
    ["Activity","#","Testimony / Impact Story"],
    ...(cmcReport.programs||[]).flatMap(p=>(p.testimonies||[]).filter(t=>t&&t.trim()).map((t,k)=>[p.name||"(Unnamed Activity)",k+1,t])),
    [],["OTHER CMC REPORTS"],
    ["Fellowship Visits:",o.fellowVisits,"","Associates as Resource Persons:",o.assocRP],
    ["Associate Visits:",o.assocVisits,"","Associates Mapped Out:",o.assocMapped],
    ["Associate Calls:",o.assocCalls,"","Students Discipled by Staff:",o.discipled],
    [],["THANK YOU FOR YOUR FAITHFULNESS"]
  ];
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"]=Array(10).fill({wch:22}).map((c,i)=>i===7?{wch:50}:c);
  return ws;
}
function buildZonalSheet(zoneName,month,year,stationGroups){
  const rows=[
    [`MONTHLY ZONAL REPORT — ${zoneName}`],["Month:",month,"Year:",year],[],
    ["FELLOWSHIP","MEMBERSHIP","ATTENDANCE","EVAN.AWARENESS","GOSPEL EXP.","NEW CONVERTS","PRAYER ENG.","INVOLVED","DAILY QT","CDF MEETINGS","CDF PRESENT","SMALL GROUP BIBLE STUDY","SCRIPTURE MEM.","STATUS"],
  ];
  stationGroups.forEach(sg=>{
    rows.push([`--- ${sg.stationName} ---`]);
    sg.fellowships.forEach(fd=>{
      if(fd.t)rows.push([fd.fellowship,fd.t.membership,fd.t.attendance,fd.t.awareness,fd.t.exposure,fd.t.converts,fd.t.prayerEng,fd.t.inPrayer,fd.t.dailyQT,fd.t.cdfMet,fd.t.cdfPresent,fd.t.sgbs,fd.t.memorization,"SUBMITTED"]);
      else rows.push([fd.fellowship,"","","","","","","","","","","","","NOT SUBMITTED"]);
    });
    const st=sumTotals(sg.fellowships.filter(f=>f.t).map(f=>f.t));
    rows.push([`${sg.stationName} SUBTOTAL`,st.membership||0,st.attendance||0,st.awareness||0,st.exposure||0,st.converts||0,st.prayerEng||0,st.inPrayer||0,st.dailyQT||0,st.cdfMet||0,st.cdfPresent||0,st.sgbs||0,st.memorization||0]);
    rows.push([]);
  });
  const all=stationGroups.flatMap(sg=>sg.fellowships);
  const gt=sumTotals(all.filter(f=>f.t).map(f=>f.t));
  rows.push(["ZONE TOTAL",gt.membership||0,gt.attendance||0,gt.awareness||0,gt.exposure||0,gt.converts||0,gt.prayerEng||0,gt.inPrayer||0,gt.dailyQT||0,gt.cdfMet||0,gt.cdfPresent||0,gt.sgbs||0,gt.memorization||0]);
  const ws=XLSX.utils.aoa_to_sheet(rows);
  ws["!cols"]=Array(14).fill({wch:18});
  return ws;
}

// ═══════════════════════════════════════════════════════
// FELLOWSHIP VIEW
// ═══════════════════════════════════════════════════════
function FellowshipView({allReports,onSave,settings,T}){
  const [report,setReport]=useState(mkReport());
  const [sec,setSec]=useState("det");
  const [week,setWeek]=useState("1");
  const [msg,setMsg]=useState({text:"",ok:true});
  const [syncing,setSyncing]=useState(false);
  const [sidebarOpen,setSidebarOpen]=useState(false);
  const [isExisting,setIsExisting]=useState(false);
  const [showPastReports,setShowPastReports]=useState(false);
  const mounted=useRef(false);
  const prevFellowship=useRef("");
  const savedSnapshot=useRef(JSON.stringify(mkReport()));

  const updHdr=useCallback((f,v)=>setReport(p=>({...p,hdr:{...p.hdr,[f]:v}})),[]);
  const updWeek=useCallback((wk,s,val)=>setReport(p=>({...p,weeks:{...p.weeks,[wk]:{...p.weeks[wk],[s]:val}}})),[]);

  const showMsg=(text,ok=true)=>{setMsg({text,ok});setTimeout(()=>setMsg({text:"",ok:true}),3000);};
  const showMsgRef=useRef(showMsg);
  showMsgRef.current=showMsg;

  // Whenever Fellowship / Month / Year change, sync the form to whatever's
  // actually saved for that combination — otherwise switching to review an
  // old report shows a blank form, and saving would silently wipe it out.
  // If nothing is saved yet for the new combination, only wipe the in-progress
  // data when the fellowship itself changed (a genuinely different report) —
  // a month/year tweak alone must never discard week data someone is mid-typing.
  useEffect(()=>{
    if(!mounted.current){
      mounted.current=true;
      prevFellowship.current=report.hdr.fellowship;
      return;
    }
    const {fellowship,month,year}=report.hdr;
    const fellowshipChanged=fellowship!==prevFellowship.current;
    prevFellowship.current=fellowship;
    if(!fellowship||fellowship==="-- Select --"){setIsExisting(false);return;}
    const key=rKey(fellowship,month,year);
    const existing=allReports[key];
    if(existing){
      setReport(existing);
      savedSnapshot.current=JSON.stringify(existing);
      setIsExisting(true);
      showMsgRef.current(`📄 Loaded your existing ${month} ${year} report for ${fellowship}.`);
    } else if(fellowshipChanged){
      setReport(p=>({...mkReport(),hdr:{...mkReport().hdr,fellowship,month,year,studentName:p.hdr.studentName,presidentName:p.hdr.presidentName}}));
      setIsExisting(false);
    } else {
      // Same fellowship, just a month/year tweak with nothing saved for it yet —
      // keep whatever's already been typed in this session.
      setIsExisting(false);
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  },[report.hdr.fellowship,report.hdr.month,report.hdr.year]);

  // Warn before leaving the tab if there are unsaved changes
  useEffect(()=>{
    const handler=(e)=>{
      if(JSON.stringify(report)!==savedSnapshot.current){
        e.preventDefault();
        e.returnValue="";
      }
    };
    window.addEventListener("beforeunload",handler);
    return ()=>window.removeEventListener("beforeunload",handler);
  },[report]);

  const handleSave=async()=>{
    if(!report.hdr.fellowship){showMsg("⚠ Select a fellowship first",false);return;}
    const key=rKey(report.hdr.fellowship,report.hdr.month,report.hdr.year);
    await onSave({...allReports,[key]:report});
    savedSnapshot.current=JSON.stringify(report);
    setIsExisting(true);

    if(settings?.scriptUrl){
      setSyncing(true);
      const res=await syncToSheets(settings.scriptUrl,report,settings.sheetName);
      setSyncing(false);
      showMsg(res.ok?"✓ Saved & synced to Google Sheets!":"✓ Saved locally (sync failed: "+res.msg+")",res.ok);
    } else {
      showMsg("✓ Report saved locally.");
    }
  };

  const exportXLSX=()=>{
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,buildFellowshipSheet(report),`${report.hdr.month} Report`);
    XLSX.writeFile(wb,`Fellowship_Report_${report.hdr.fellowship||"Unknown"}_${report.hdr.month}_${report.hdr.year}.xlsx`);
  };

  const wkData=report.weeks[week];
  const T_val=totals(report.weeks);
  const wk_val=weekStats(wkData);
  const sections=SECS(T);
  const filledWeeks=WEEKS.filter(w=>Object.values(report.weeks[w].det).some(Boolean));

  const sidebarBg=T.mode==="bright"?"#1E3A6E":"rgba(0,0,0,0.25)";
  const sidebarText=T.mode==="bright"?"rgba(255,255,255,0.7)":"";

  return <div style={{display:"flex",height:"100%",overflow:"hidden",position:"relative"}}>

    {/* Mobile toggle button for this sidebar */}
    <button className="fv-sidebar-toggle" onClick={()=>setSidebarOpen(o=>!o)} style={{
      display:"none",position:"absolute",top:10,left:10,zIndex:40,
      background:sidebarBg,color:"#fff",border:`1px solid ${T.border}`,borderRadius:8,
      padding:"7px 10px",cursor:"pointer",fontSize:15,fontFamily:"inherit"}}>
      {sidebarOpen?"✕":"☰"}
    </button>

    {/* Overlay behind the sidebar on mobile */}
    {sidebarOpen&&<div className="fv-sidebar-overlay" onClick={()=>setSidebarOpen(false)} style={{display:"none",position:"absolute",inset:0,background:"rgba(0,0,0,0.45)",zIndex:20}}/>}

    {/* Sidebar */}
    <div className={"fv-sidebar"+(sidebarOpen?" open":"")} style={{width:200,flexShrink:0,background:sidebarBg,borderRight:`1px solid ${T.border}`,display:"flex",flexDirection:"column",overflow:"hidden"}}>
      <div style={{padding:"14px 10px",flex:1,overflowY:"auto"}}>
        <div style={{fontSize:9.5,color:T.mode==="bright"?"rgba(255,255,255,0.5)":T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8,padding:"0 4px"}}>Sections</div>
        {sections.map(s=>{
          const active=sec===s.id;
          const textCol=T.mode==="bright"?(active?"#fff":"rgba(255,255,255,0.65)"):( active?s.color:T.muted);
          const bg=T.mode==="bright"?(active?"rgba(255,255,255,0.15)":"transparent"):(active?`${s.color}18`:"transparent");
          return <button key={s.id} onClick={()=>{setSec(s.id);setSidebarOpen(false);}} style={{
            display:"flex",alignItems:"center",gap:8,width:"100%",textAlign:"left",
            padding:"8px 10px",borderRadius:7,marginBottom:2,cursor:"pointer",border:"none",
            background:bg,color:textCol,fontSize:12.5,fontWeight:active?700:400,
            borderLeft:active?`3px solid ${T.mode==="bright"?"#fff":s.color}`:"3px solid transparent",
            fontFamily:"inherit",transition:"all .15s"}}>
            <span style={{fontSize:14}}>{s.icon}</span>{s.label}
          </button>;
        })}
        <div style={{height:1,background:T.mode==="bright"?"rgba(255,255,255,0.2)":T.border,margin:"12px 0"}}/>
        <div style={{fontSize:9.5,color:T.mode==="bright"?"rgba(255,255,255,0.5)":T.muted,textTransform:"uppercase",letterSpacing:"1px",marginBottom:8,padding:"0 4px"}}>Reporting Week</div>
        {WEEKS.map(w=>{
          const active=week===w;
          const filled=filledWeeks.includes(w);
          return <button key={w} onClick={()=>{setWeek(w);setSidebarOpen(false);}} style={{
            display:"flex",alignItems:"center",gap:8,width:"100%",textAlign:"left",
            padding:"7px 10px",borderRadius:7,marginBottom:2,cursor:"pointer",border:"none",
            background:T.mode==="bright"?(active?"rgba(255,255,255,0.2)":"transparent"):(active?`${T.green}18`:"transparent"),
            color:T.mode==="bright"?(active?"#fff":"rgba(255,255,255,0.65)"):(active?T.green:T.muted),
            fontSize:12.5,fontWeight:active?700:400,fontFamily:"inherit",transition:"all .15s"}}>
            <span style={{width:7,height:7,borderRadius:"50%",flexShrink:0,
              background:filled?(T.mode==="bright"?"#86EFAC":T.green):(T.mode==="bright"?"rgba(255,255,255,0.25)":T.sub)}}/>
            Week {w}
          </button>;
        })}
      </div>
    </div>

    {/* Main */}
    <div className="app-main" style={{flex:1,overflow:"auto",padding:"18px 22px",background:T.bg}}>
      {/* Header card */}
      <div className="app-card" style={{background:T.mode==="bright"?"linear-gradient(135deg,#EBF5FF,#DBEAFE)":"linear-gradient(135deg,rgba(22,163,74,0.12),rgba(202,138,4,0.08))",
        border:`1px solid ${T.mode==="bright"?"#BFDBFE":T.border}`,borderRadius:14,padding:"16px 18px",marginBottom:14,
        boxShadow:T.mode==="bright"?"0 2px 8px rgba(0,0,0,0.06)":"none"}}>
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12,flexWrap:"wrap",gap:8}}>
          <span style={{fontSize:11,fontWeight:800,color:T.mode==="bright"?"#1E3A6E":T.green,textTransform:"uppercase",letterSpacing:"1px"}}>Report Header</span>
          {report.hdr.fellowship&&report.hdr.fellowship!=="-- Select --"&&(
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              {isExisting
                ?<span style={{fontSize:10.5,fontWeight:700,color:T.mode==="bright"?"#1D4ED8":T.disc,background:T.mode==="bright"?"#DBEAFE":T.discBg,padding:"3px 10px",borderRadius:20}}>📄 Editing Existing Report</span>
                :<span style={{fontSize:10.5,fontWeight:700,color:T.mode==="bright"?"#15803D":T.green,background:T.mode==="bright"?"#DCFCE7":T.glow,padding:"3px 10px",borderRadius:20}}>✨ New Report</span>}
              <button onClick={()=>setShowPastReports(true)} style={{fontSize:10.5,fontWeight:700,color:T.prog,background:T.progBg,border:`1px solid ${T.prog}44`,padding:"3px 10px",borderRadius:20,cursor:"pointer",fontFamily:"inherit"}}>📚 Past Reports</button>
            </div>
          )}
        </div>
        <Grid cols="1fr 1fr 1fr">
          <Sel label="Fellowship" v={report.hdr.fellowship} set={v=>updHdr("fellowship",v)} opts={["-- Select --",...FELLOWSHIPS]} T={T}/>
          <TIn label="Student Name" v={report.hdr.studentName} set={v=>updHdr("studentName",v)} ph="Your full name" T={T}/>
          <TIn label="Fellowship President" v={report.hdr.presidentName} set={v=>updHdr("presidentName",v)} ph="President's name" T={T}/>
        </Grid>
        <Grid cols="1fr 1fr 1fr 1fr 1fr">
          <Sel label="Month" v={report.hdr.month} set={v=>updHdr("month",v)} opts={MONTHS} T={T}/>
          <Sel label="Year" v={report.hdr.year} set={v=>updHdr("year",v)} opts={YEARS} T={T}/>
          <Sel label="Reporting Week" v={week} set={setWeek} opts={WEEKS} T={T}/>
          <Field label="Submission Date" T={T}><input type="date" value={report.hdr.subDate} onChange={e=>updHdr("subDate",e.target.value)} style={inpStyle(T)}/></Field>
          <TIn label="Special Program this Week" v={report.hdr.specialProg} set={v=>updHdr("specialProg",v)} ph="If any…" T={T}/>
        </Grid>
      </div>

      {["det","wit","cdf","pry"].includes(sec)&&(
        <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
          <span style={{fontSize:11.5,color:T.muted}}>Editing data for:</span>
          <span style={{background:T.mode==="bright"?"#DBEAFE":T.glow,color:T.mode==="bright"?"#1D4ED8":T.green,padding:"3px 12px",borderRadius:20,fontSize:12,fontWeight:700,border:`1px solid ${T.mode==="bright"?"#93C5FD":"transparent"}`}}>Week {week}</span>
          <span style={{fontSize:11,color:T.sub}}>— change it above in the header, or in the sidebar</span>
        </div>
      )}

      {sec==="det"  &&<DetSection  d={wkData.det}  set={v=>updWeek(week,"det",v)}  T={T}/>}
      {sec==="wit"  &&<WitSection  d={wkData.wit}  set={v=>updWeek(week,"wit",v)}  T={T}/>}
      {sec==="cdf"  &&<CdfSection  d={wkData.cdf}  set={v=>updWeek(week,"cdf",v)}  T={T}/>}
      {sec==="pry"  &&<PrySection  d={wkData.pry}  set={v=>updWeek(week,"pry",v)}  T={T}/>}
      {sec==="lead" &&<LeadSection d={report.lead} set={v=>setReport(p=>({...p,lead:v}))} T={T}/>}
      {sec==="prog" &&<ProgSection zonalProg={report.zonalProg} majorProg={report.majorProg} programs={report.programs}
        setZonal={v=>setReport(p=>({...p,zonalProg:v}))} setMajor={v=>setReport(p=>({...p,majorProg:v}))} setProgs={v=>setReport(p=>({...p,programs:v}))} T={T}/>}
      {sec==="oth"  &&<OthSection  d={report.other} set={v=>setReport(p=>({...p,other:v}))} isCMC={false} T={T}/>}
      {sec==="wsum" &&<>
        <SecHead title={`Weekly Summary — Week ${week} (Auto-Computed)`} color={T.green} icon="📊" T={T}/>
        <Grid cols="repeat(4,1fr)" gap={8}>
          {[["Membership",wk_val.membership,T.oth],["Total Attendance",wk_val.attendance,T.oth],["Gospel Exposure",wk_val.exposure,T.wit],["New Converts",wk_val.converts,T.wit]].map(([l,v,c])=><StatCard key={l} label={l} value={v} color={c} T={T}/>)}
        </Grid>
        <div style={{height:8}}/>
        <Grid cols="repeat(4,1fr)" gap={8}>
          {[["Prayer Engagements",wk_val.prayerEng,T.pry],["Involved in Prayer",wk_val.inPrayer,T.pry],["Daily Quiet Time",wk_val.dailyQT,T.pry],["CDF Meetings",wk_val.cdfMet,T.disc]].map(([l,v,c])=><StatCard key={l} label={l} value={v} color={c} T={T}/>)}
        </Grid>
        <div style={{height:8}}/>
        <Grid cols="repeat(4,1fr)" gap={8}>
          {[["CDF Members Present",wk_val.cdfPresent,T.disc],["Scripture Trained",wk_val.studsTrained,T.pry],["Small Group Bible Study",wk_val.sgbs,T.pry],["Scripture Memorization",wk_val.memorization,T.pry]].map(([l,v,c])=><StatCard key={l} label={l} value={v} color={c} T={T}/>)}
        </Grid>
      </>}
      {sec==="msum" &&<>
        <SecHead title="Monthly Summary — All Weeks Combined (Auto-Computed)" color={T.green} icon="🗓" T={T}/>
        <div style={{fontSize:11.5,color:T.muted,marginBottom:12}}>Collates every reporting week entered so far for this fellowship into one monthly total.</div>
        <Grid cols="repeat(4,1fr)" gap={8}>
          {[["Membership",T_val.membership,T.oth],["Total Attendance",T_val.attendance,T.oth],["Gospel Exposure",T_val.exposure,T.wit],["New Converts",T_val.converts,T.wit]].map(([l,v,c])=><StatCard key={l} label={l} value={v} color={c} T={T}/>)}
        </Grid>
        <div style={{height:8}}/>
        <Grid cols="repeat(4,1fr)" gap={8}>
          {[["Prayer Engagements",T_val.prayerEng,T.pry],["Involved in Prayer",T_val.inPrayer,T.pry],["Daily Quiet Time",T_val.dailyQT,T.pry],["CDF Meetings",T_val.cdfMet,T.disc]].map(([l,v,c])=><StatCard key={l} label={l} value={v} color={c} T={T}/>)}
        </Grid>
        <div style={{height:8}}/>
        <Grid cols="repeat(4,1fr)" gap={8}>
          {[["CDF Members Present",T_val.cdfPresent,T.disc],["Scripture Trained",T_val.studsTrained,T.pry],["Small Group Bible Study",T_val.sgbs,T.pry],["Scripture Memorization",T_val.memorization,T.pry]].map(([l,v,c])=><StatCard key={l} label={l} value={v} color={c} T={T}/>)}
        </Grid>
      </>}

      {/* Action bar */}
      <div style={{display:"flex",alignItems:"center",gap:10,marginTop:18,paddingTop:16,borderTop:`1px solid ${T.border}`,justifyContent:"flex-end",flexWrap:"wrap"}}>
        {msg.text&&<span style={{fontSize:12.5,fontWeight:600,color:msg.ok?(T.mode==="bright"?"#15803D":T.green):"#DC2626"}}>{msg.text}</span>}
        {syncing&&<span style={{fontSize:12,color:T.muted,fontStyle:"italic"}}>Syncing…</span>}
        {sec==="oth"&&(
          <button onClick={async()=>{await handleSave();setSec("wsum");}} style={{display:"flex",alignItems:"center",gap:6,background:T.mode==="bright"?"#16A34A":T.glow,border:`1.5px solid ${T.green}`,color:T.mode==="bright"?"#fff":T.green,padding:"10px 22px",borderRadius:9,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
            ✅ Submit and View Summary
          </button>
        )}
        {!["oth","wsum","msum"].includes(sec)&&(
          <button onClick={async()=>{await handleSave();const order=sections.map(s=>s.id);setSec(order[order.indexOf(sec)+1]);}} style={{display:"flex",alignItems:"center",gap:6,background:T.progBg,border:`1.5px solid ${T.prog}55`,color:T.prog,padding:"10px 22px",borderRadius:9,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
            Next Section →
          </button>
        )}
        <button onClick={handleSave} style={{display:"flex",alignItems:"center",gap:6,background:T.mode==="bright"?"#1E3A6E":T.glow,border:`1.5px solid ${T.mode==="bright"?"#1E3A6E":T.green}`,color:T.mode==="bright"?"#fff":T.green,padding:"10px 22px",borderRadius:9,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>
          {settings?.scriptUrl?"💾 Save & Sync to Sheets":"💾 Save Report"}
        </button>
        {["wsum","msum"].includes(sec)&&<>
          <button onClick={exportXLSX} style={{display:"flex",alignItems:"center",gap:6,background:T.discBg,border:`1.5px solid ${T.disc}55`,color:T.disc,padding:"10px 18px",borderRadius:9,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>📊 Export Excel</button>
          <button onClick={()=>window.print()} style={{display:"flex",alignItems:"center",gap:6,background:T.progBg,border:`1.5px solid ${T.prog}55`,color:T.prog,padding:"10px 18px",borderRadius:9,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>🖨 Print / PDF</button>
        </>}
      </div>
    </div>
    {showPastReports&&<PastReportsModal fellowship={report.hdr.fellowship} allReports={allReports}
      onPick={(month,year)=>{setReport(p=>({...p,hdr:{...p.hdr,month,year}}));}}
      onClose={()=>setShowPastReports(false)} T={T}/>}
  </div>;
}

// ═══════════════════════════════════════════════════════
// CMC VIEW
// ═══════════════════════════════════════════════════════
function CMCView({allReports,onSave,settings,T}){
  const [cmc,setCMC]=useState(mkCMC());
  // Arbitrary, freely-editable list of fellowships this CMC report covers.
  // Defaults to no preselection — the user builds the list themselves,
  // optionally jump-starting it from a Zone/Station preset below.
  const [coords,setCoords]=useState([]);
  const [zone,setZone]=useState("South East Zone");
  const [station,setStation]=useState(ZONES["South East Zone"][0].name);
  const [msg,setMsg]=useState("");
  const [fSearch,setFSearch]=useState("");

  const updHdr=(f,v)=>setCMC(p=>({...p,hdr:{...p.hdr,[f]:v}}));
  const {month,year}=cmc.hdr;

  const toggleFellowship=(f)=>setCoords(p=>p.includes(f)?p.filter(x=>x!==f):[...p,f]);
  const loadPreset=()=>{
    const stObj=ZONES[zone]?.find(s=>s.name===station);
    if(!stObj)return;
    setCoords(p=>Array.from(new Set([...p,...stObj.fellowships])));
  };
  const clearAll=()=>setCoords([]);
  const selectAll=()=>setCoords(FELLOWSHIPS.slice());
  const visibleFellowships=FELLOWSHIPS.filter(f=>f.toLowerCase().includes(fSearch.toLowerCase()));

  const fdList=coords.map(f=>{
    const r=allReports[rKey(f,month,year)];
    return{fellowship:f,t:r?totals(r.weeks):null};
  });
  const grand=sumTotals(fdList.filter(f=>f.t).map(f=>f.t));
  const submitted=fdList.filter(f=>f.t).length;

  const exportXLSX=()=>{
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,buildCMCSheet(cmc,fdList),`${month} CMC Report`);
    XLSX.writeFile(wb,`CMC_Report_${cmc.hdr.staffName||"CMC"}_${month}_${year}.xlsx`);
  };

  const th={textAlign:"left",padding:"8px",color:T.muted,borderBottom:`1px solid ${T.border}`,fontWeight:600,fontSize:10,textTransform:"uppercase",letterSpacing:"0.4px",whiteSpace:"nowrap"};

  return <div className="app-main" style={{overflow:"auto",padding:"18px 22px",background:T.bg}}>
    <Card T={T} accent={T.mode==="bright"?"#1E3A6E":undefined}>
      <div style={{fontSize:11,fontWeight:800,color:T.mode==="bright"?"#1E3A6E":T.disc,textTransform:"uppercase",letterSpacing:"1px",marginBottom:12}}>CMC Staff Information</div>
      <Grid cols="1fr 1fr 1fr 1fr">
        <TIn label="Staff Name" v={cmc.hdr.staffName} set={v=>updHdr("staffName",v)} ph="Full name" T={T}/>
        <TIn label="Station" v={cmc.hdr.station} set={v=>updHdr("station",v)} ph="e.g. Accra - ZMC" T={T}/>
        <Sel label="Devotional Life" v={cmc.hdr.devLife} set={v=>updHdr("devLife",v)} opts={DEVLIFE} T={T}/>
        <TIn label="Books Reading this Week" v={cmc.hdr.books} set={v=>updHdr("books",v)} ph="Book title" T={T}/>
      </Grid>
      <Grid cols="1fr 1fr 1fr">
        <Sel label="Month" v={cmc.hdr.month} set={v=>updHdr("month",v)} opts={MONTHS} T={T}/>
        <Sel label="Year" v={cmc.hdr.year} set={v=>updHdr("year",v)} opts={YEARS} T={T}/>
        <Field label="Submission Date" T={T}><input type="date" value={cmc.hdr.subDate} onChange={e=>updHdr("subDate",e.target.value)} style={inpStyle(T)}/></Field>
      </Grid>
    </Card>

    <Card T={T} accent={T.prog}>
      <div style={{fontSize:11,fontWeight:800,color:T.prog,textTransform:"uppercase",letterSpacing:"1px",marginBottom:10}}>Fellowships Coordinated by this CMC</div>
      <div style={{fontSize:11.5,color:T.muted,marginBottom:12,lineHeight:1.5}}>
        Tick any combination of fellowships below — a CMC isn't limited to a fixed station list. Optionally use a Zone/Station preset to jump-start the selection, then add or remove individual fellowships freely.
      </div>

      {/* Optional preset loader */}
      <div style={{display:"flex",gap:8,alignItems:"flex-end",flexWrap:"wrap",marginBottom:14,padding:"10px 12px",background:T.mode==="bright"?"#F8FAFC":"rgba(255,255,255,0.03)",borderRadius:10,border:`1px dashed ${T.border}`}}>
        <div style={{flex:"1 1 140px"}}><Sel label="Zone" v={zone} set={z=>{setZone(z);setStation(ZONES[z][0].name);}} opts={Object.keys(ZONES)} mb={0} T={T}/></div>
        <div style={{flex:"1 1 160px"}}><Sel label="Sub-Zone" v={station} set={setStation} opts={(ZONES[zone]||[]).map(s=>s.name)} mb={0} T={T}/></div>
        <button onClick={loadPreset} style={{background:T.progBg,border:`1.5px solid ${T.prog}55`,color:T.prog,padding:"9px 14px",borderRadius:8,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>+ Add Sub-Zone Group</button>
        <button onClick={selectAll} style={{background:"transparent",border:`1.5px solid ${T.border}`,color:T.muted,padding:"9px 14px",borderRadius:8,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>Select All</button>
        <button onClick={clearAll} style={{background:"transparent",border:`1.5px solid ${T.border}`,color:T.muted,padding:"9px 14px",borderRadius:8,fontSize:12,fontWeight:700,cursor:"pointer",fontFamily:"inherit",whiteSpace:"nowrap"}}>Clear</button>
      </div>

      {/* Search + checklist */}
      <input type="text" value={fSearch} onChange={e=>setFSearch(e.target.value)} placeholder="Search fellowships…"
        style={inpStyle(T,{marginBottom:10})}/>
      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(190px,1fr))",gap:6,maxHeight:220,overflowY:"auto",padding:"2px 2px 4px"}}>
        {visibleFellowships.map(f=>{
          const checked=coords.includes(f);
          return <label key={f} onClick={()=>toggleFellowship(f)} style={{
            display:"flex",alignItems:"center",gap:8,cursor:"pointer",
            background:checked?(T.mode==="bright"?"#EAF7EF":T.glow):T.inputBg,
            border:`1px solid ${checked?T.green:T.border}`,borderRadius:8,
            padding:"7px 10px",fontSize:12.5,color:checked?T.green:T.text,fontWeight:checked?700:500}}>
            <input type="checkbox" checked={checked} onChange={()=>toggleFellowship(f)} style={{accentColor:T.green,cursor:"pointer"}}/>
            {f}
          </label>;
        })}
        {visibleFellowships.length===0&&<div style={{gridColumn:"1/-1",fontSize:12,color:T.muted,padding:"8px 2px"}}>No fellowships match "{fSearch}".</div>}
      </div>

      <div style={{fontSize:12,color:T.muted,marginTop:12}}>
        Coordinating <strong style={{color:T.text}}>{coords.length}</strong> fellowship(s)
        {coords.length>0&&<>: <span style={{color:T.prog}}>{coords.join(", ")}</span></>}
        {coords.length>0&&<>{" · "}<span style={{fontWeight:700,color:submitted===coords.length?T.green:T.wit}}>{submitted}/{coords.length} submitted</span></>}
      </div>
    </Card>

    <Grid cols="repeat(4,1fr)" gap={8} style={{marginBottom:14}}>
      {[["Total Mbsp",grand.membership,T.oth],["Gospel Exp.",grand.exposure,T.wit],["New Converts",grand.converts,T.wit],["Prayer Eng.",grand.prayerEng,T.pry]].map(([l,v,c])=><StatCard key={l} label={l} value={v} color={c} T={T}/>)}
    </Grid>

    <Card T={T} accent={T.disc}>
      <div style={{fontSize:12,fontWeight:800,color:T.disc,textTransform:"uppercase",letterSpacing:"0.8px",marginBottom:12}}>Fellowship Monthly Data — {month} {year}</div>
      {coords.length===0?(
        <div style={{padding:"18px 8px",color:T.muted,fontSize:12.5,fontStyle:"italic"}}>No fellowships selected yet — tick some above to pull in their {month} {year} data.</div>
      ):
      <div style={{overflowX:"auto"}}>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
          <thead><tr>{["Fellowship","Mbsp","Attd","Aware","Gospel Exp.","Converts","Prayer","Daily QT","CDF Mtgs","SGBS","Status",""].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
          <tbody>
            {fdList.map(fd=><tr key={fd.fellowship} style={{borderBottom:`1px solid ${T.border}22`}}>
              <td style={{padding:"9px 8px",color:T.text,fontWeight:600}}>{fd.fellowship}</td>
              {fd.t?(<>
                {[fd.t.membership,fd.t.attendance].map((v,i)=><td key={i} style={{padding:"9px 8px",color:T.text}}>{v||0}</td>)}
                {[fd.t.awareness,fd.t.exposure,fd.t.converts].map((v,i)=><td key={i} style={{padding:"9px 8px",color:T.wit,fontWeight:600}}>{v||0}</td>)}
                {[fd.t.prayerEng,fd.t.dailyQT].map((v,i)=><td key={i} style={{padding:"9px 8px",color:T.pry}}>{v||0}</td>)}
                <td style={{padding:"9px 8px",color:T.disc}}>{fd.t.cdfMet||0}</td>
                <td style={{padding:"9px 8px",color:T.pry}}>{fd.t.sgbs||0}</td>
                <td style={{padding:"9px 8px"}}><span style={{background:`${T.green}22`,color:T.green,padding:"2px 8px",borderRadius:4,fontSize:10,fontWeight:700}}>✓</span></td>
              </>):<td colSpan={9} style={{padding:"9px 8px",color:"#EF4444",fontSize:11,fontStyle:"italic"}}>No report for {month} {year}</td>}
              <td style={{padding:"9px 8px"}}>
                <button onClick={()=>toggleFellowship(fd.fellowship)} title="Remove from this CMC report"
                  style={{background:"transparent",border:"none",color:T.muted,cursor:"pointer",fontSize:14,fontWeight:700,padding:"2px 6px"}}>×</button>
              </td>
            </tr>)}
            <tr style={{borderTop:`2px solid ${T.disc}44`}}>
              <td style={{padding:"10px 8px",color:T.disc,fontWeight:800}}>TOTAL</td>
              {[grand.membership,grand.attendance,grand.awareness,grand.exposure,grand.converts,grand.prayerEng,grand.dailyQT,grand.cdfMet,grand.sgbs].map((v,i)=><td key={i} style={{padding:"10px 8px",fontWeight:700,color:i>=2&&i<=4?T.wit:i>=5?T.pry:T.text}}>{v||0}</td>)}
              <td/><td/>
            </tr>
          </tbody>
        </table>
      </div>}
    </Card>

    <LeadSection d={cmc.lead} set={v=>setCMC(p=>({...p,lead:v}))} T={T}/>
    <ProgSection zonalProg={cmc.zonalProg} majorProg={cmc.majorProg} programs={cmc.programs}
      setZonal={v=>setCMC(p=>({...p,zonalProg:v}))} setMajor={v=>setCMC(p=>({...p,majorProg:v}))} setProgs={v=>setCMC(p=>({...p,programs:v}))} T={T}/>
    <OthSection d={cmc.other} set={v=>setCMC(p=>({...p,other:v}))} isCMC={true} T={T}/>

    <div style={{display:"flex",gap:8,marginTop:16,paddingTop:16,borderTop:`1px solid ${T.border}`,justifyContent:"flex-end"}}>
      {msg&&<span style={{fontSize:12.5,color:T.green,fontWeight:600}}>{msg}</span>}
      <button onClick={exportXLSX} style={{display:"flex",alignItems:"center",gap:6,background:T.mode==="bright"?"#1E3A6E":T.glow,border:`1.5px solid ${T.mode==="bright"?"#1E3A6E":T.green}`,color:T.mode==="bright"?"#fff":T.green,padding:"10px 20px",borderRadius:9,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>📊 Export Excel</button>
      <button onClick={()=>window.print()} style={{display:"flex",alignItems:"center",gap:6,background:T.progBg,border:`1.5px solid ${T.prog}55`,color:T.prog,padding:"10px 20px",borderRadius:9,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>🖨 Print / PDF</button>
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════
// ZONAL VIEW
// ═══════════════════════════════════════════════════════
function ZonalView({allReports,T}){
  const [zone,setZone]=useState("South East Zone");
  const [month,setMonth]=useState(MONTHS[new Date().getMonth()]);
  const [year,setYear]=useState(String(new Date().getFullYear()));

  const stations=ZONES[zone]||[];
  const stationGroups=stations.map(st=>({
    stationName:st.name,
    fellowships:st.fellowships.map(f=>{const r=allReports[rKey(f,month,year)];return{fellowship:f,t:r?totals(r.weeks):null};})
  }));
  const allFds=stationGroups.flatMap(sg=>sg.fellowships);
  const grand=sumTotals(allFds.filter(f=>f.t).map(f=>f.t));
  const submitted=allFds.filter(f=>f.t).length;
  const total=allFds.length;

  const exportXLSX=()=>{
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,buildZonalSheet(zone,month,year,stationGroups),`${month} Zonal`);
    XLSX.writeFile(wb,`Zonal_Report_${zone.replace(/\s+/g,"_")}_${month}_${year}.xlsx`);
  };

  const th={textAlign:"left",padding:"6px 7px",color:T.muted,borderBottom:`1px solid ${T.border}`,fontWeight:600,fontSize:9.5,textTransform:"uppercase",letterSpacing:"0.4px",whiteSpace:"nowrap"};

  return <div className="app-main" style={{overflow:"auto",padding:"18px 22px",background:T.bg}}>
    <div style={{background:T.mode==="bright"?"linear-gradient(135deg,#1E3A6E,#1D4ED8)":"linear-gradient(135deg,rgba(6,182,212,0.12),rgba(22,163,74,0.08))",border:`1px solid ${T.mode==="bright"?"#1E3A6E":T.border}`,borderRadius:14,padding:"16px 18px",marginBottom:14}}>
      <div style={{fontSize:11,fontWeight:800,color:T.mode==="bright"?"rgba(255,255,255,0.8)":T.prog,textTransform:"uppercase",letterSpacing:"1px",marginBottom:12}}>Zone Consolidation View</div>
      <Grid cols="1fr 1fr 1fr">
        <Sel label="Zone" v={zone} set={setZone} opts={Object.keys(ZONES)} T={T.mode==="bright"?{...T,inputBg:"rgba(255,255,255,0.12)",selectBg:"rgba(255,255,255,0.12)",border:"rgba(255,255,255,0.3)",text:"#fff",muted:"rgba(255,255,255,0.7)"}:T}/>
        <Sel label="Month" v={month} set={setMonth} opts={MONTHS} T={T.mode==="bright"?{...T,inputBg:"rgba(255,255,255,0.12)",selectBg:"rgba(255,255,255,0.12)",border:"rgba(255,255,255,0.3)",text:"#fff",muted:"rgba(255,255,255,0.7)"}:T}/>
        <Sel label="Year" v={year} set={setYear} opts={YEARS} T={T.mode==="bright"?{...T,inputBg:"rgba(255,255,255,0.12)",selectBg:"rgba(255,255,255,0.12)",border:"rgba(255,255,255,0.3)",text:"#fff",muted:"rgba(255,255,255,0.7)"}:T}/>
      </Grid>
      <div style={{fontSize:12,color:T.mode==="bright"?"rgba(255,255,255,0.75)":T.muted,marginTop:4}}>
        <span style={{color:T.mode==="bright"?"#86EFAC":T.green,fontWeight:700}}>{submitted}</span> of <strong style={{color:T.mode==="bright"?"#fff":T.text}}>{total}</strong> fellowships have submitted reports for {month} {year}
        <span style={{marginLeft:14,color:submitted===total?(T.mode==="bright"?"#86EFAC":T.green):"#FCD34D",fontWeight:600}}>{submitted===total?"✓ All submitted":submitted===0?"⚠ No submissions yet":`⚠ ${total-submitted} pending`}</span>
      </div>
    </div>

    <Grid cols="repeat(4,1fr)" gap={8} style={{marginBottom:8}}>
      {[["Membership",grand.membership,T.oth],["Attendance",grand.attendance,T.oth],["Gospel Exposure",grand.exposure,T.wit],["New Converts",grand.converts,T.wit]].map(([l,v,c])=><StatCard key={l} label={l} value={v} color={c} T={T}/>)}
    </Grid>
    <Grid cols="repeat(4,1fr)" gap={8} style={{marginBottom:16}}>
      {[["Prayer Eng.",grand.prayerEng,T.pry],["Daily QT",grand.dailyQT,T.pry],["CDF Meetings",grand.cdfMet,T.disc],["SGBS",grand.sgbs,T.pry]].map(([l,v,c])=><StatCard key={l} label={l} value={v} color={c} T={T}/>)}
    </Grid>

    {stationGroups.map(sg=>{
      const st=sumTotals(sg.fellowships.filter(f=>f.t).map(f=>f.t));
      const sub=sg.fellowships.filter(f=>f.t).length;
      return <Card key={sg.stationName} T={T} accent={T.prog} style={{marginBottom:12}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:10,flexWrap:"wrap",gap:8}}>
          <div>
            <div style={{fontSize:12,fontWeight:800,color:T.prog,textTransform:"uppercase",letterSpacing:"0.6px"}}>{sg.stationName}</div>
            <div style={{fontSize:11,color:T.muted,marginTop:2}}>{sub}/{sg.fellowships.length} submitted</div>
          </div>
          <div style={{display:"flex",gap:10,fontSize:11.5,flexWrap:"wrap"}}>
            {[["Mbsp",st.membership,T.muted],["Attd",st.attendance,T.muted],["Exp",st.exposure,T.wit],["Conv",st.converts,T.wit]].map(([l,v,c])=>(
              <span key={l} style={{color:T.muted}}>{l}: <strong style={{color:c}}>{v||0}</strong></span>
            ))}
          </div>
        </div>
        <div style={{overflowX:"auto"}}>
          <table style={{width:"100%",borderCollapse:"collapse",fontSize:11.5}}>
            <thead><tr>{["Fellowship","Mbsp","Attd","Aware","Exp.","Conv.","Prayer","QT","CDF","SGBS","Mem.","Status"].map(h=><th key={h} style={th}>{h}</th>)}</tr></thead>
            <tbody>
              {sg.fellowships.map(fd=><tr key={fd.fellowship} style={{borderBottom:`1px solid ${T.border}22`}}>
                <td style={{padding:"8px 7px",color:T.text,fontWeight:600}}>{fd.fellowship}</td>
                {fd.t?(<>
                  {[fd.t.membership,fd.t.attendance].map((v,i)=><td key={i} style={{padding:"8px 7px",color:T.text}}>{v||0}</td>)}
                  {[fd.t.awareness,fd.t.exposure,fd.t.converts].map((v,i)=><td key={i} style={{padding:"8px 7px",color:T.wit,fontWeight:600}}>{v||0}</td>)}
                  {[fd.t.prayerEng,fd.t.dailyQT,fd.t.cdfMet,fd.t.sgbs,fd.t.memorization].map((v,i)=><td key={i} style={{padding:"8px 7px",color:T.pry}}>{v||0}</td>)}
                  <td style={{padding:"8px 7px"}}><span style={{background:`${T.green}22`,color:T.green,padding:"2px 7px",borderRadius:4,fontSize:9.5,fontWeight:700}}>✓</span></td>
                </>):<td colSpan={11} style={{padding:"8px 7px",color:"#EF444488",fontSize:10,fontStyle:"italic"}}>No report — {month}</td>}
              </tr>)}
            </tbody>
          </table>
        </div>
      </Card>;
    })}

    <div style={{display:"flex",gap:8,marginTop:8,justifyContent:"flex-end"}}>
      <button onClick={exportXLSX} style={{display:"flex",alignItems:"center",gap:6,background:T.mode==="bright"?"#1E3A6E":T.glow,border:`1.5px solid ${T.mode==="bright"?"#1E3A6E":T.green}`,color:T.mode==="bright"?"#fff":T.green,padding:"10px 20px",borderRadius:9,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>📊 Export Zonal Excel</button>
      <button onClick={()=>window.print()} style={{display:"flex",alignItems:"center",gap:6,background:T.progBg,border:`1.5px solid ${T.prog}55`,color:T.prog,padding:"10px 20px",borderRadius:9,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>🖨 Print / PDF</button>
    </div>
  </div>;
}

// ═══════════════════════════════════════════════════════
// MAIN APP
// ═══════════════════════════════════════════════════════
const VIEWS=[
  {id:"fellowship",label:"Fellowship Report",emoji:"📝",desc:"Weekly Reports by Campus Fellowships"},
];
const GATED_VIEWS=[
  {id:"cmc",       label:"CMC Report",       emoji:"📊",desc:"Campus Ministry Coordinator aggregated view"},
  {id:"zonal",     label:"Zonal Report",     emoji:"🌐",desc:"Zone-wide consolidated monthly report"},
];
const ALL_VIEWS=[...VIEWS,...GATED_VIEWS];

function RestrictedView({label,onUnlock,T}){
  return <div style={{flex:1,display:"flex",alignItems:"center",justifyContent:"center",padding:24}}>
    <div style={{textAlign:"center",maxWidth:360}}>
      <div style={{fontSize:40,marginBottom:10}}>🔒</div>
      <div style={{fontSize:15,fontWeight:800,color:T.text,marginBottom:6}}>{label} is restricted</div>
      <div style={{fontSize:12.5,color:T.muted,marginBottom:16,lineHeight:1.5}}>This area is only for authorised staff. Enter the password for {label} to continue.</div>
      <button onClick={onUnlock} style={{background:T.mode==="bright"?"#1E3A6E":T.glow,border:`1.5px solid ${T.mode==="bright"?"#1E3A6E":T.green}`,color:T.mode==="bright"?"#fff":T.green,padding:"10px 22px",borderRadius:9,fontSize:13,fontWeight:700,cursor:"pointer",fontFamily:"inherit"}}>🔒 Staff Access</button>
    </div>
  </div>;
}

export default function App(){
  const [view,setView]         = useState("fellowship");
  const [reps,setReps]         = useState({});
  const [settings,setSettings] = useState({});
  const [ready,setReady]       = useState(false);
  const [themeKey,setTheme]    = useState("bright");
  const [showSettings,setShowSettings] = useState(false);
  const [navOpen,setNavOpen]   = useState(false);
  const [unlocked,setUnlocked] = useState(loadUnlocked);
  const [showStaffGate,setShowStaffGate] = useState(false);

  const handleUnlock=(role)=>setUnlocked(p=>{const n={...p,[role]:true};saveUnlocked(n);return n;});
  const goToRole=(role)=>{
    setShowStaffGate(false);
    if(role==="sheets") setShowSettings(true);
    else setView(role);
  };

  const T = themeKey==="bright" ? BRIGHT : DARK;

  useEffect(()=>{
    Promise.all([loadData(),loadSettings()]).then(([d,s])=>{setReps(d);setSettings(s);setReady(true);});
  },[]);

  const handleSave    = useCallback(async nr=>{setReps(nr);await saveData(nr);},[]);
  const handleSaveSettings = useCallback(async s=>{setSettings(s);await saveSettings(s);setShowSettings(false);},[]);

  const importInputRef = useRef(null);
  const [backupMsg,setBackupMsg] = useState("");
  const exportBackup=()=>{
    const blob=new Blob([JSON.stringify(reps,null,2)],{type:"application/json"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;
    a.download=`GHAFES_Backup_${new Date().toISOString().split("T")[0]}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };
  const importBackup=(e)=>{
    const file=e.target.files?.[0];
    if(!file)return;
    const reader=new FileReader();
    reader.onload=async(ev)=>{
      try{
        const imported=JSON.parse(ev.target.result);
        if(typeof imported!=="object"||Array.isArray(imported))throw new Error("bad shape");
        const merged={...reps,...imported};
        setReps(merged);
        await saveData(merged);
        setBackupMsg(`✓ Restored ${Object.keys(imported).length} report(s).`);
      }catch{
        setBackupMsg("⚠ That file doesn't look like a valid GHAFES backup.");
      }
      setTimeout(()=>setBackupMsg(""),4000);
    };
    reader.readAsText(file);
    e.target.value="";
  };

  const repCount=Object.keys(reps).length;

  const navBg    = T.mode==="bright"?"#1E3A6E":T.navBg;
  const navText  = T.mode==="bright"?"#fff":T.text;
  const navMuted = T.mode==="bright"?"rgba(255,255,255,0.6)":T.muted;

  return <div className="app-shell" style={{display:"flex",flexDirection:"column",background:T.bg,fontFamily:"system-ui,-apple-system,'Segoe UI',Roboto,sans-serif",color:T.text,overflow:"hidden"}}>

    {/* TOP BAR */}
    <div style={{background:navBg,borderBottom:`1px solid ${T.mode==="bright"?"#1E3A6E":T.border}`,flexShrink:0,zIndex:50}}>
      <div className="topbar-row" style={{display:"flex",alignItems:"stretch"}}>

        {/* Logo */}
        <div className="topbar-logo" style={{display:"flex",alignItems:"center",gap:10,padding:"8px 18px",borderRight:`1px solid ${T.mode==="bright"?"rgba(255,255,255,0.2)":T.border}`,minWidth:180}}>
          <img src={LOGO} alt="GHAFES Logo" style={{width:40,height:40,objectFit:"contain",borderRadius:4,flexShrink:0}}/>
          <div>
            <div style={{fontSize:13,fontWeight:800,color:"#fff",letterSpacing:"-0.3px",lineHeight:1}}>GHAFES</div>
            <div className="topbar-tagline" style={{fontSize:10,color:T.mode==="bright"?"rgba(255,255,255,0.65)":"#7DB992",letterSpacing:"0.4px"}}>Ministry Reports</div>
          </div>
        </div>

        {/* Nav tabs — Fellowship is always public; CMC/Zonal only appear once unlocked this session */}
        {[...VIEWS,...GATED_VIEWS.filter(v=>unlocked[v.id])].map(v=>{
          const active=view===v.id;
          return <button key={v.id} onClick={()=>setView(v.id)} className="nav-tab" title={v.label} style={{
            display:"flex",alignItems:"center",gap:8,padding:"0 20px",cursor:"pointer",border:"none",
            background:"transparent",
            borderBottom:active?`3px solid ${T.mode==="bright"?"#56C8E8":"#16A34A"}`:"3px solid transparent",
            color:active?"#fff":navMuted,
            fontSize:13,fontWeight:active?700:500,fontFamily:"inherit",transition:"all .15s"}}>
            <span style={{fontSize:16}}>{v.emoji}</span>
            <span className="nav-tab-label">{v.label}</span>
          </button>;
        })}

        {/* Staff Access — the only way in for anyone not yet unlocked */}
        <button onClick={()=>setShowStaffGate(true)} className="nav-tab" title="Staff Access (CMC / Zonal / Sheets)" style={{
          display:"flex",alignItems:"center",gap:8,padding:"0 20px",cursor:"pointer",border:"none",
          background:"transparent",borderBottom:"3px solid transparent",color:navMuted,
          fontSize:13,fontWeight:500,fontFamily:"inherit",transition:"all .15s"}}>
          <span style={{fontSize:16}}>🔒</span>
          <span className="nav-tab-label">Staff Access</span>
        </button>

        {/* Right controls */}
        <div style={{marginLeft:"auto",display:"flex",alignItems:"center",gap:8,padding:"0 16px"}}>
          <div className="topbar-badges" style={{display:"flex",gap:6,fontSize:10}}>
            {[[T.wit,"Witness"],[T.disc,"Discipleship"],[T.lead,"Leadership"]].map(([c,l])=>(
              <span key={l} style={{background:`${c}22`,color:c,padding:"3px 8px",borderRadius:4,fontWeight:700,textTransform:"uppercase",letterSpacing:"0.5px",border:`1px solid ${c}44`}}>{l}</span>
            ))}
          </div>
          <span className="topbar-savedcount" style={{fontSize:11,color:navMuted,marginLeft:4}}>
            <span style={{color:T.mode==="bright"?"#86EFAC":T.greenLt,fontWeight:700}}>{repCount}</span> saved
          </span>
          {backupMsg&&<span style={{fontSize:10.5,fontWeight:600,color:backupMsg.startsWith("✓")?(T.mode==="bright"?"#86EFAC":T.greenLt):"#FCA5A5"}}>{backupMsg}</span>}
          <button onClick={exportBackup} title="Download a backup of all saved reports on this device" style={{
            background:"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.2)",color:"#fff",
            borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:13,fontFamily:"inherit",
            display:"flex",alignItems:"center",gap:5}}>
            ⬇ <span className="topbar-btn-label" style={{fontSize:11,fontWeight:600}}>Backup</span>
          </button>
          <input ref={importInputRef} type="file" accept="application/json" onChange={importBackup} style={{display:"none"}}/>
          <button onClick={()=>importInputRef.current?.click()} title="Restore reports from a backup file" style={{
            background:"rgba(255,255,255,0.1)",border:"1px solid rgba(255,255,255,0.2)",color:"#fff",
            borderRadius:8,padding:"6px 10px",cursor:"pointer",fontSize:13,fontFamily:"inherit",
            display:"flex",alignItems:"center",gap:5}}>
            ⬆ <span className="topbar-btn-label" style={{fontSize:11,fontWeight:600}}>Restore</span>
          </button>
          {unlocked.sheets&&<button onClick={()=>setShowSettings(true)} title="Google Sheets Settings" style={{
            background:settings?.scriptUrl?(T.mode==="bright"?"rgba(134,239,172,0.2)":T.glow):"rgba(255,255,255,0.1)",
            border:`1px solid ${settings?.scriptUrl?"rgba(134,239,172,0.5)":"rgba(255,255,255,0.2)"}`,
            color:settings?.scriptUrl?"#86EFAC":"#fff",borderRadius:8,padding:"6px 12px",cursor:"pointer",
            fontSize:14,fontFamily:"inherit",display:"flex",alignItems:"center",gap:5}}>
            {settings?.scriptUrl?"🔗":"⚙"} <span className="topbar-btn-label" style={{fontSize:11,fontWeight:600}}>{settings?.scriptUrl?"Sheets ✓":"Sheets"}</span>
          </button>}
          <button onClick={()=>setTheme(t=>t==="dark"?"bright":"dark")} title="Toggle theme" style={{
            background:"rgba(255,255,255,0.12)",border:"1px solid rgba(255,255,255,0.2)",
            color:"#fff",borderRadius:8,padding:"6px 12px",cursor:"pointer",fontSize:14,fontFamily:"inherit",
            display:"flex",alignItems:"center",gap:5}}>
            {themeKey==="dark"?"☀":"🌙"} <span className="topbar-btn-label" style={{fontSize:11,fontWeight:600}}>{themeKey==="dark"?"Light":"Dark"}</span>
          </button>
        </div>
      </div>

      {/* Sub-header */}
      <div className="desktop-subheader" style={{padding:"5px 18px 7px",borderTop:`1px solid ${T.mode==="bright"?"rgba(255,255,255,0.15)":T.border+"33"}`,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
        <span style={{fontSize:11,color:navMuted}}>{ALL_VIEWS.find(v=>v.id===view)?.desc}</span>
        <span style={{fontSize:10,color:navMuted,fontStyle:"italic"}}>{T.mode==="bright"?"☀ GHAFES Brand Theme":"🌙 Dark Theme"}</span>
      </div>
    </div>

    {/* CONTENT */}
    <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>
      {!ready?(
        <div style={{display:"flex",alignItems:"center",justifyContent:"center",flex:1,flexDirection:"column",gap:14}}>
          <img src={LOGO} alt="GHAFES" style={{width:64,height:64,objectFit:"contain",opacity:0.8}}/>
          <div style={{fontSize:14,color:T.muted}}>Loading saved reports…</div>
        </div>
      ):(
        <div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}>
          {view==="fellowship"&&<div style={{flex:1,overflow:"hidden",display:"flex",flexDirection:"column"}}><FellowshipView allReports={reps} onSave={handleSave} settings={settings} T={T}/></div>}
          {view==="cmc"&&(unlocked.cmc
            ?<CMCView allReports={reps} onSave={handleSave} settings={settings} T={T}/>
            :<RestrictedView label="CMC Report" onUnlock={()=>setShowStaffGate(true)} T={T}/>)}
          {view==="zonal"&&(unlocked.zonal
            ?<ZonalView allReports={reps} T={T}/>
            :<RestrictedView label="Zonal Report" onUnlock={()=>setShowStaffGate(true)} T={T}/>)}
        </div>
      )}
    </div>

    {/* FOOTER */}
    <div style={{flexShrink:0,background:navBg,borderTop:`1px solid ${T.mode==="bright"?"rgba(255,255,255,0.15)":T.border}`,padding:"7px 18px",textAlign:"center"}}>
      <span style={{fontFamily:"'Cormorant Garamond',Georgia,serif",fontStyle:"italic",fontWeight:600,fontSize:13.5,letterSpacing:"0.3px",color:T.mode==="bright"?"rgba(255,255,255,0.85)":"#9FD8B5"}}>
        Knowing Christ and Making Him Known!
      </span>
    </div>

    {showStaffGate&&<StaffGateModal unlocked={unlocked} onUnlock={handleUnlock} onClose={()=>setShowStaffGate(false)} onGoTo={goToRole} T={T}/>}
    {showSettings&&unlocked.sheets&&<SettingsModal settings={settings} onSave={handleSaveSettings} onClose={()=>setShowSettings(false)} T={T}/>}

    <style>{`
      * { box-sizing:border-box; }
      .app-shell { height:100vh; height:100dvh; }
      input[type=number]::-webkit-inner-spin-button { opacity:0.4; }
      ::-webkit-scrollbar { width:6px; height:6px; }
      ::-webkit-scrollbar-track { background:transparent; }
      ::-webkit-scrollbar-thumb { background:${T.scrollThumb}; border-radius:3px; }
      select option { background:${T.mode==="bright"?"#EBF5FF":"#0D1525"}; color:${T.text}; }
      @media print {
        body { background:white!important; }
        * { color:#000!important; background:transparent!important; border-color:#ccc!important; }
      }

      /* ── Mobile: header stays, but shrinks to icon-only tabs and hides secondary chrome ── */
      @media (max-width: 720px) {
        .topbar-row { flex-wrap:wrap; }
        .topbar-logo { min-width:0!important; padding:6px 10px!important; }
        .topbar-tagline { display:none; }
        .nav-tab { padding:0 12px!important; }
        .nav-tab-label { display:none; }
        .topbar-badges, .topbar-savedcount, .topbar-btn-label, .desktop-subheader { display:none!important; }

        .fv-sidebar-toggle { display:flex!important; align-items:center; justify-content:center; }
        .fv-sidebar-overlay { display:block!important; }
        .fv-sidebar {
          position:absolute; top:0; left:0; bottom:0; width:200px;
          transform:translateX(-100%); transition:transform .2s ease; z-index:30;
        }
        .fv-sidebar.open { transform:translateX(0); }

        /* Content padding — reclaim space on narrow screens */
        .app-main { padding:12px 12px calc(20px + env(safe-area-inset-bottom, 0px))!important; }
        .app-card { padding:12px 13px!important; }

        /* Grids collapse to 2 columns on tablets/large phones, 1 on narrow phones */
        .app-grid { grid-template-columns:repeat(2,1fr)!important; gap:8px!important; }
      }

      @media (max-width: 480px) {
        .app-grid { grid-template-columns:1fr!important; }
      }
    `}</style>
  </div>;
}
