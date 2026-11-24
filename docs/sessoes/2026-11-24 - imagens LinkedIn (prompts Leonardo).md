# 24/11/2026 — Imagens para o LinkedIn (cibersegurança) e prompts para o Leonardo.ai

## Pedido
A partir de 6 imagens de referência (tema segurança digital / proteção de dados), criar:
1. uma imagem **futurista**, juntando o melhor de cada referência;
2. uma imagem **psicodélica**, com o mesmo contexto;
3. para cada uma, dois prompts para o Leonardo.ai: um **detalhado** e um no estilo `/imagine prompt: ... —c 10 —ar`.

## O que veio de cada referência
| Referência | Elemento aproveitado |
|---|---|
| 1 (mão + balança) | mão tocando o holograma; balança da justiça (compliance/LGPD) |
| 2 (cadeado azul) | cadeado grande com circuitos e engrenagens; grade de blueprint |
| 3 (mão + cadeado) | anéis HUD concêntricos, feixes de luz, iluminação dramática |
| 4 (cadeado + circuitos) | trilhas de circuito simétricas com nós; código binário |
| 5 (escudo no chip) | escudo; cidade de chips/servidores com bokeh |
| 6 (notebook) | ícones de painel de dados (barras, gráficos) nos anéis |

## Entregas
- Imagens geradas pelo Gamma (`generate_image`, 16:9), links na conversa e na biblioteca de mídia do Gamma.
- Prompts abaixo.

---

## 1. Imagem FUTURISTA

### Prompt detalhado
Cinematic photorealistic cybersecurity and data-protection key visual, 16:9. On the left, a human hand emerging from
deep darkness, index finger extended, about to touch a giant glowing holographic interface floating in mid-air. The
interface is a set of concentric luminous HUD rings in electric blue and cyan, slightly tilted in 3D perspective, with
segmented arcs, tick marks, small rectangular data bars, tiny dashboard icons (bar charts, gears, check marks) and thin
light streaks radiating outward. At the exact center of the rings, a bold glowing padlock with a keyhole; inside the
padlock body a subtle translucent shield outline and a small scales-of-justice icon, symbolizing compliance and data
protection law (LGPD). Behind the rings, a dark navy background with a faint blueprint grid, glowing circuit-board
traces fanning out symmetrically left and right ending in bright nodes, drifting streams of binary code 0 and 1, and
semi-transparent rotating gears. In the lower foreground, out of focus, a futuristic microchip city of glowing server
towers with bokeh. Color palette: deep black, midnight navy, electric blue, cyan, tiny accents of white light. Dramatic
rim light on the hand, volumetric light, shallow depth of field, ultra sharp details, 8k, premium corporate technology
aesthetic, clean negative space on the right side for text. No text, no letters, no logos.

Negative prompt (Leonardo): text, letters, watermark, logo, extra fingers, deformed hand, blurry, low quality,
oversaturated, cartoon.

Sugestão de ajustes no Leonardo: modelo Leonardo Phoenix ou Lightning XL, estilo "Cinematic", proporção 16:9
(post) ou 4:1 (capa do perfil), Alchemy ligado.

### Prompt no estilo /imagine
/imagine prompt: cinematic photorealistic image of a human hand emerging from darkness, index finger about to touch a giant glowing holographic HUD made of concentric electric blue and cyan rings, at the center a bold glowing padlock with keyhole containing a translucent shield and a small scales of justice icon, symmetrical glowing circuit board traces with bright nodes radiating left and right, streams of binary code 0 and 1, faint blueprint grid and semi-transparent gears in the background, out of focus futuristic microchip city of server towers with bokeh in the foreground, deep black and midnight navy palette, volumetric light, dramatic rim light, ultra detailed, 8k, premium corporate cybersecurity aesthetic, negative space on the right, no text —c 10 —ar 16:9

---

## 2. Imagem PSICODÉLICA

### Prompt detalhado
Psychedelic visionary art version of a cybersecurity key visual, 16:9. A human hand reaching from the left, index
finger touching a giant swirling kaleidoscopic mandala made of concentric holographic rings. At the center of the
mandala, a glowing iridescent padlock with a keyhole that opens into a fractal tunnel; inside the padlock a shimmering
shield and a small scales-of-justice symbol. The rings are made of liquid neon circuit-board traces, flowing binary code
0 and 1 melting like paint, interlocking gears turning into flower petals, sacred geometry patterns and fractal
spirals, with tiny dashboard icons dissolving into the pattern. Circuit lines radiate outward in waves like an aura,
ending in glowing orbs. In the lower foreground, a dreamlike city of microchip server towers dissolving into liquid
light. Color palette: ultraviolet, magenta, hot pink, acid green, electric cyan and gold, high saturation, chromatic
aberration, glowing rim light, 1970s psychedelic poster meets modern digital art, visionary art style, ultra detailed,
symmetrical composition, 8k. No text, no letters, no logos.

Negative prompt (Leonardo): text, letters, watermark, logo, extra fingers, deformed hand, muddy colors, dull, low
contrast, blurry.

Sugestão de ajustes no Leonardo: modelo Leonardo Phoenix ou Leonardo Kino XL, estilo "Vibrant" ou "Illustration",
proporção 16:9, Alchemy ligado.

### Prompt no estilo /imagine
/imagine prompt: psychedelic visionary art of a human hand touching a giant swirling kaleidoscopic mandala of holographic rings, at the center a glowing iridescent padlock whose keyhole opens into a fractal tunnel, with a shimmering shield and a small scales of justice symbol, rings made of liquid neon circuit board traces, melting binary code 0 and 1, gears turning into flower petals, sacred geometry and fractal spirals, circuit lines radiating like an aura ending in glowing orbs, dreamlike microchip city dissolving into liquid light, ultraviolet magenta hot pink acid green cyan and gold, high saturation, chromatic aberration, 1970s psychedelic poster meets modern digital art, symmetrical, ultra detailed, 8k, no text —c 10 —ar 16:9

---

## Resultado no Gamma
- Arquivos locais (fora do git): `imagens LinkedIn\linkedin_futurista.jpg` e `linkedin_psicodelica.jpg` (2752 × 1536).
- Futurista (tipo `photo`): saiu como pedido na 1ª tentativa.
- Psicodélica: com o tipo `abstract`, o Gamma ignorou o assunto e gerou só ondas roxas. Refeita com o tipo
  `illustration` e "MUST clearly show…" no prompt: saiu com todos os elementos, mas em traço chapado (cartoon) e com
  duas mãos. Para um visual psicodélico realista, usar o prompt acima no Leonardo.
- Créditos do Gamma: 3 × 70 = 210 gastos, restam 55.

## Observações
- `—c 10` (chaos) é parâmetro do Midjourney; o Leonardo ignora. Lá, a proporção se escolhe no painel, e a variação
  vem do número de imagens por geração.
- Capa do perfil do LinkedIn: 1584 × 396 (4:1). Post: 1200 × 627 (≈ 1,91:1) ou 1:1. Trocar o `—ar` conforme o uso.
- Nada de dados da Autopel nesta sessão.
