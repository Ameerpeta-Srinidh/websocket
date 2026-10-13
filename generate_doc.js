
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  HeadingLevel, AlignmentType, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageNumber, PageBreak, TabStopType, TabStopPosition
} = require('docx');
const fs = require('fs');

// ── Helpers ───────────────────────────────────────────────────────────────────

const FONT = "Calibri";
const CODE_FONT = "Courier New";
const ACCENT = "1F4E79";   // dark blue
const ACCENT2 = "2E75B6";  // mid blue
const GREEN  = "375623";
const ORANGE = "7F3F00";
const GRAY_BG = "F2F2F2";
const BLUE_BG = "DEEAF1";
const GREEN_BG = "E2EFDA";
const ORANGE_BG = "FCE4D6";
const PURPLE_BG = "EAE0F0";

const border = { style: BorderStyle.SINGLE, size: 1, color: "CCCCCC" };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 320, after: 160 },
    children: [new TextRun({ text, font: FONT, size: 32, bold: true, color: ACCENT })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 240, after: 100 },
    children: [new TextRun({ text, font: FONT, size: 26, bold: true, color: ACCENT2 })]
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 180, after: 80 },
    children: [new TextRun({ text, font: FONT, size: 22, bold: true, color: "333333" })]
  });
}
function p(text, opts = {}) {
  return new Paragraph({
    spacing: { before: 60, after: 80 },
    children: [new TextRun({ text, font: FONT, size: 22, ...opts })]
  });
}
function pRuns(runs) {
  return new Paragraph({
    spacing: { before: 60, after: 80 },
    children: runs.map(r =>
      new TextRun({ font: r.code ? CODE_FONT : FONT, size: r.code ? 20 : 22, ...r })
    )
  });
}
function blank() {
  return new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun("")] });
}
function rule() {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: "2E75B6", space: 1 } },
    children: [new TextRun("")]
  });
}
function bullet(text, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { before: 40, after: 40 },
    children: [new TextRun({ text, font: FONT, size: 22 })]
  });
}
function bulletRuns(runs, level = 0) {
  return new Paragraph({
    numbering: { reference: "bullets", level },
    spacing: { before: 40, after: 40 },
    children: runs.map(r => new TextRun({ font: r.code ? CODE_FONT : FONT, size: r.code ? 20 : 22, ...r }))
  });
}

function codeBlock(lines, bg = GRAY_BG) {
  const rows = lines.map(line =>
    new TableRow({
      children: [new TableCell({
        borders: noBorders,
        width: { size: 9360, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 120, right: 120 },
        shading: { fill: bg, type: ShadingType.CLEAR },
        children: [new Paragraph({
          spacing: { before: 0, after: 0 },
          children: [new TextRun({ text: line || " ", font: CODE_FONT, size: 18, color: "1A1A1A" })]
        })]
      })]
    })
  );
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    margins: { top: 80, bottom: 80 },
    rows
  });
}

function noteBox(title, text, bg, titleColor) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    margins: { top: 100, bottom: 100 },
    rows: [
      new TableRow({ children: [new TableCell({
        borders,
        width: { size: 9360, type: WidthType.DXA },
        margins: { top: 100, bottom: 100, left: 180, right: 180 },
        shading: { fill: bg, type: ShadingType.CLEAR },
        children: [
          new Paragraph({ spacing: { before: 0, after: 60 }, children: [
            new TextRun({ text: title, font: FONT, size: 22, bold: true, color: titleColor })
          ]}),
          new Paragraph({ spacing: { before: 0, after: 0 }, children: [
            new TextRun({ text, font: FONT, size: 21, color: "333333" })
          ]})
        ]
      })]})
    ]
  });
}

function twoCol(left, right, leftPct = 50) {
  const total = 9360;
  const lw = Math.round(total * leftPct / 100);
  const rw = total - lw;
  return new Table({
    width: { size: total, type: WidthType.DXA },
    columnWidths: [lw, rw],
    rows: [new TableRow({ children: [
      new TableCell({ borders: noBorders, width: { size: lw, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 0, right: 180 },
        children: left }),
      new TableCell({ borders: noBorders, width: { size: rw, type: WidthType.DXA },
        margins: { top: 0, bottom: 0, left: 180, right: 0 },
        children: right })
    ]})]
  });
}

function sectionTag(text, color = ACCENT) {
  return new Paragraph({
    spacing: { before: 200, after: 60 },
    children: [new TextRun({
      text: `  ${text}  `,
      font: FONT, size: 18, bold: true, color: "FFFFFF",
      highlight: undefined,
      shading: { fill: color, type: ShadingType.CLEAR }
    })]
  });
}

// ── Document content ──────────────────────────────────────────────────────────

const children = [

  // ── COVER ─────────────────────────────────────────────────────────────────
  new Paragraph({ spacing: { before: 1200, after: 100 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "DS-CMGAN", font: FONT, size: 72, bold: true, color: ACCENT })] }),
  new Paragraph({ spacing: { before: 0, after: 60 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "A Self-Teaching Guide to Writing the Phonetic Modules", font: FONT, size: 30, color: ACCENT2 })] }),
  new Paragraph({ spacing: { before: 40, after: 40 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Dysarthria-aware CMGAN — Code-it-yourself edition", font: FONT, size: 24, color: "555555", italics: true })] }),
  blank(), blank(),
  new Paragraph({ spacing: { before: 0, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "For: Srinidh  ·  Advisor: Prof. Hemant A. Patil", font: FONT, size: 22, color: "777777" })] }),
  new Paragraph({ spacing: { before: 0, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: "Speech Research Lab, DA-IICT  ·  Interspeech 2026", font: FONT, size: 22, color: "777777" })] }),
  blank(), blank(),

  // How to use this document
  noteBox(
    "How to use this document",
    "Each chapter teaches you one concept — the syntax, the math, and the reason it exists in " +
    "DS-CMGAN. At the end of each chapter there is a 'Write It Yourself' exercise. " +
    "Do not look at any reference code until you have attempted the exercise. " +
    "The goal is that after reading this document you can write every module from scratch.",
    BLUE_BG, ACCENT
  ),

  new Paragraph({ children: [new PageBreak()] }),

  // ── CHAPTER 0: PyTorch Foundations ────────────────────────────────────────
  h1("Chapter 0 — PyTorch Foundations You Must Know"),
  p("Before writing any DS-CMGAN code, you need to own these four concepts. Everything else builds on them."),
  blank(),

  h2("0.1  nn.Module — the base class for every neural network layer"),
  p("Every class you write for DS-CMGAN inherits from nn.Module. This gives you three things automatically:"),
  bullet("Parameter tracking: any nn.Linear, nn.Conv1d, etc. you define inside __init__ is automatically registered as a trainable parameter."),
  bullet("Device handling: .to(device) moves all parameters to GPU/CPU in one call."),
  bullet("State dict: .state_dict() and .load_state_dict() for saving/loading checkpoints."),
  blank(),
  p("The skeleton every module must follow:"),
  codeBlock([
    "import torch.nn as nn",
    "",
    "class MyModule(nn.Module):          # always inherit from nn.Module",
    "    def __init__(self, ...):        # define all layers here",
    "        super().__init__()          # REQUIRED — calls nn.Module.__init__",
    "        self.layer1 = nn.Linear(...)  # registered automatically",
    "",
    "    def forward(self, x):           # defines the computation",
    "        return self.layer1(x)       # PyTorch calls this when you do: m(x)"
  ]),

  blank(),
  noteBox("Rule", "Never put computation in __init__. Never put layer definitions in forward. __init__ defines structure. forward defines flow.", GREEN_BG, GREEN),
  blank(),

  h2("0.2  How tensors work — shape is everything"),
  p("Every operation in DS-CMGAN is a tensor shape transformation. If you know the input shape and the output shape you want, you can write the operation."),
  blank(),
  p("Key shape operations you will use constantly:"),
  codeBlock([
    "x = torch.randn(4, 320, 100, 64)   # [B, T, F', C]",
    "",
    "# Permute: reorder dimensions",
    "x.permute(0, 2, 1, 3)              # → [B, F', T, C]",
    "",
    "# Reshape / view: merge or split dimensions",
    "x.reshape(4 * 100, 320, 64)       # → [BF', T, C]  fold F' into batch",
    "x.view(4, 320, 100, 64)            # → back to [B, T, F', C]",
    "",
    "# Transpose: swap exactly two dimensions",
    "x.transpose(1, 2)                  # swap dim1 and dim2",
    "",
    "# Unsqueeze / expand: add and broadcast dimensions",
    "p = torch.randn(4, 320, 64)        # [B, T, C]",
    "p.unsqueeze(1)                     # → [B, 1, T, C]",
    "p.unsqueeze(1).expand(-1,100,-1,-1)  # → [B, F', T, C]  (-1 = keep)",
    "",
    "# Mean over a dimension",
    "x.mean(dim=2)                      # → [B, T, C]  average over F'"
  ]),
  blank(),

  h2("0.3  requires_grad and torch.no_grad() — controlling the gradient tape"),
  p("PyTorch builds a computation graph as your data flows forward. Backward() walks this graph to compute gradients. Two mechanisms let you control which parts are included:"),
  blank(),
  codeBlock([
    "# 1. requires_grad = False on a parameter",
    "#    That parameter is excluded from the gradient computation permanently.",
    "#    Its weights will never be updated by the optimizer.",
    "for param in model.parameters():",
    "    param.requires_grad = False",
    "",
    "# 2. torch.no_grad() context manager",
    "#    No computation graph is built inside this block.",
    "#    Used for inference, validation, and frozen forward passes.",
    "#    Saves memory — no graph = no intermediate activations stored.",
    "with torch.no_grad():",
    "    output = frozen_model(input)   # graph not built",
    "",
    "# The difference:",
    "# requires_grad=False  →  permanent, parameter-level",
    "# torch.no_grad()      →  temporary, block-level, for any tensor"
  ]),
  blank(),
  noteBox(
    "Why both are needed for Wav2Vec2",
    "requires_grad=False on Wav2Vec2 parameters tells the optimizer to never update them. " +
    "torch.no_grad() inside forward() prevents building a computation graph through Wav2Vec2's " +
    "50+ transformer layers, saving ~40% GPU memory per forward pass. You need both.",
    ORANGE_BG, ORANGE
  ),
  blank(),

  h2("0.4  register_buffer — tensors that travel with the model but are not trained"),
  p("Some tensors need to be saved with the model (move to GPU with .to(device), saved in checkpoints) but should NOT be trained. Example: the formant weight vector in your loss function."),
  codeBlock([
    "class MyLoss(nn.Module):",
    "    def __init__(self):",
    "        super().__init__()",
    "        W = torch.tensor([1.0, 2.0, 1.5, ...])   # computed, not learned",
    "        self.register_buffer('W', W)                # NOT self.W = nn.Parameter(W)",
    "",
    "    def forward(self, x, y):",
    "        return ((x - y)**2 * self.W).mean()        # W is on same device as x"
  ]),
  blank(),
  p("Comparison:"),
  codeBlock([
    "self.W = nn.Parameter(tensor)    # trainable — optimizer updates this",
    "self.register_buffer('W', tensor) # NOT trainable — travels with model, device-aware",
    "self.W = tensor                  # NOT part of model — won't move to GPU, not saved"
  ]),

  new Paragraph({ children: [new PageBreak()] }),

  // ── CHAPTER 1: Frozen Phonetic Extractor ─────────────────────────────────
  h1("Chapter 1 — Frozen Phonetic Feature Extractor"),
  blank(),
  noteBox(
    "What this module does",
    "CMGAN's conformer processes the spectrogram but doesn't know which phoneme is being produced " +
    "at each frame. This module gives it that knowledge by running a frozen Wav2Vec2 on the same " +
    "waveform and producing a 64-dim phoneme feature vector per STFT frame. " +
    "Output shape: [B, T_stft, 64].",
    BLUE_BG, ACCENT
  ),
  blank(),

  h2("1.1  What you need to import"),
  codeBlock([
    "import torch",
    "import torch.nn as nn",
    "import torch.nn.functional as F",
    "from transformers import Wav2Vec2Model   # pip install transformers"
  ]),
  blank(),

  h2("1.2  Loading a pre-trained model from HuggingFace"),
  p("HuggingFace Transformers lets you load any pre-trained model in one line:"),
  codeBlock([
    "model = Wav2Vec2Model.from_pretrained('facebook/wav2vec2-base-960h')",
    "# This downloads ~360MB on first run, then caches it.",
    "# 'facebook/wav2vec2-base-960h' = trained on 960 hours of LibriSpeech.",
    "",
    "# Access a specific layer's output:",
    "output = model(waveform)            # waveform: [B, L]",
    "hidden = output.last_hidden_state   # [B, T_w2v, 768]",
    "#                                          ↑",
    "#                               768-dim representation per frame (~50 fps at 16kHz)"
  ]),
  blank(),

  h2("1.3  The temporal alignment problem"),
  p("Wav2Vec2 produces ~50 frames/second. CMGAN's STFT (hop=100 samples, 16kHz) produces 160 frames/second. You must stretch Wav2Vec2's output to match CMGAN."),
  blank(),
  p("Solution: F.interpolate — 1D linear interpolation along the time axis."),
  codeBlock([
    "# F.interpolate needs shape [B, Channels, Time]",
    "hidden = hidden.transpose(1, 2)          # [B, 768, T_w2v]  ← channels second",
    "",
    "aligned = F.interpolate(",
    "    hidden,",
    "    size=T_stft,                          # target number of frames",
    "    mode='linear',                        # linear interpolation",
    "    align_corners=False                   # standard choice",
    ")                                         # → [B, 768, T_stft]",
    "",
    "aligned = aligned.transpose(1, 2)        # [B, T_stft, 768]  ← back to normal"
  ]),
  blank(),

  h2("1.4  The projection layer"),
  p("Wav2Vec2 outputs 768 dimensions. CMGAN uses C=64 channels throughout. A single trainable Linear layer bridges these:"),
  codeBlock([
    "self.proj = nn.Linear(768, 64)   # this IS trained — learns to compress",
    "                                  # phoneme representations into CMGAN space",
    "",
    "# In forward():",
    "out = self.proj(aligned)         # [B, T_stft, 768] → [B, T_stft, 64]"
  ]),
  blank(),

  h2("1.5  Freezing — the full pattern"),
  codeBlock([
    "# After loading Wav2Vec2, freeze it immediately:",
    "for param in self.wav2vec2.parameters():",
    "    param.requires_grad = False    # never update these weights",
    "self.wav2vec2.eval()               # disable dropout inside Wav2Vec2",
    "",
    "# In forward(), use no_grad for the frozen forward pass:",
    "with torch.no_grad():",
    "    hidden = self.wav2vec2(waveform).last_hidden_state"
  ]),
  blank(),

  noteBox(
    "Write It Yourself — Exercise 1",
    "Using only this chapter, write the complete FrozenPhoneticExtractor class. " +
    "It should have: (a) __init__ that loads Wav2Vec2, freezes it, and creates self.proj. " +
    "(b) forward(self, waveform, T_stft) that returns [B, T_stft, 64]. " +
    "Do not look at any reference code. Attempt it fully before checking.",
    PURPLE_BG, "5B2380"
  ),

  new Paragraph({ children: [new PageBreak()] }),

  // ── CHAPTER 2: Multi-Head Attention ───────────────────────────────────────
  h1("Chapter 2 — Multi-Head Attention: From Scratch to Cross-Attention"),
  blank(),
  noteBox(
    "What you will write",
    "PP-CA (Phoneme Posterior Cross-Attention): after each TS-Conformer block, " +
    "speech features H query phoneme features P. Each speech frame asks: what phoneme " +
    "is happening here, and what correction does that imply? Output shape same as input: [B, T, F', C].",
    BLUE_BG, ACCENT
  ),
  blank(),

  h2("2.1  Self-attention vs cross-attention"),
  p("Self-attention: Q, K, V all come from the same source X."),
  p("Cross-attention: Q comes from source A, K and V come from source B."),
  blank(),
  codeBlock([
    "# Self-attention (standard MHSA in CMGAN's conformer):",
    "Q = X @ W_Q    # query: what am I looking for?",
    "K = X @ W_K    # key:   what do I have to offer?",
    "V = X @ W_V    # value: what is the actual content?",
    "",
    "# Cross-attention (PP-CA — your module):",
    "Q = H @ W_Q    # query: what does speech feature H look for?",
    "K = P @ W_K    # key:   what does phoneme feature P have to offer?",
    "V = P @ W_V    # value: what phoneme content should be extracted?"
  ]),
  blank(),

  h2("2.2  The full attention equation"),
  p("Given Q, K, V matrices:"),
  codeBlock([
    "import math",
    "d_k = 64 // 4    # dimension per head = d_model / n_heads = 16",
    "",
    "scores = torch.bmm(Q, K.transpose(-2, -1)) / math.sqrt(d_k)",
    "#  Q: [batch, T, d_k]   K: [batch, T, d_k]",
    "#  scores: [batch, T, T]  — how much each frame attends to each other frame",
    "",
    "attn = torch.softmax(scores, dim=-1)   # rows sum to 1",
    "Z    = torch.bmm(attn, V)              # weighted combination of values"
  ]),
  blank(),

  h2("2.3  Multi-head: splitting and merging"),
  p("Multiple heads allow the model to attend to different aspects simultaneously. Each head gets a slice of d_model."),
  codeBlock([
    "n_heads = 4",
    "d_model = 64",
    "d_k     = d_model // n_heads   # = 16",
    "",
    "# Input: [BFp, T, d_model]  (BFp = batch × F' after reshaping)",
    "# Split into heads: [BFp, T, d_model] → [BFp, T, n_heads, d_k] → [BFp, n_heads, T, d_k]",
    "def split_heads(x, BFp, T):",
    "    return x.view(BFp, T, n_heads, d_k).transpose(1, 2)",
    "    # view: reshape last dim 64 → 4 heads × 16",
    "    # transpose: put head dim before time dim",
    "",
    "# After attention computation: merge heads back",
    "# [BFp, n_heads, T, d_k] → [BFp, T, d_model]",
    "Z = Z.transpose(1, 2)              # [BFp, T, n_heads, d_k]",
    "Z = Z.reshape(BFp, T, d_model)    # [BFp, T, 64]  ← merge"
  ]),
  blank(),

  h2("2.4  The 4D reshape for PP-CA"),
  p("CMGAN's features have shape [B, T, F', C]. PP-CA operates along the time axis. To do this with standard attention (which operates on 3D tensors), fold F' into the batch:"),
  codeBlock([
    "B, T, Fp, C = H.shape    # e.g. (4, 320, 100, 64)",
    "",
    "# Fold F' into batch dimension",
    "Ht = H.permute(0, 2, 1, 3)       # [B, F', T, C]",
    "Ht = Ht.reshape(B * Fp, T, C)    # [BFp, T, C]  — now a 3D tensor",
    "",
    "# Broadcast P (phoneme features, [B, T, C]) across Fp",
    "Pt = P.unsqueeze(1)                                         # [B, 1, T, C]",
    "Pt = Pt.expand(-1, Fp, -1, -1)                              # [B, Fp, T, C]",
    "Pt = Pt.reshape(B * Fp, T, C)                               # [BFp, T, C]",
    "",
    "# Now: Ht and Pt are both [BFp, T, C] — ready for cross-attention",
    "",
    "# After attention, restore original shape:",
    "out = out.reshape(B, Fp, T, C)                              # unfold",
    "out = out.permute(0, 2, 1, 3)                               # [B, T, Fp, C]"
  ]),
  blank(),

  h2("2.5  Residual connection and LayerNorm"),
  p("Every attention block in CMGAN ends with: output = LayerNorm(input + dropout(Z)). This is the residual connection."),
  codeBlock([
    "self.norm = nn.LayerNorm(d_model)    # normalises along last dimension",
    "self.drop = nn.Dropout(0.1)",
    "",
    "# In forward():",
    "out = self.norm(Ht + self.drop(Z))  # residual: add input back to output"
  ]),
  blank(),

  h2("2.6  Defining the weight matrices"),
  p("W_Q, W_K, W_V, W_O are all nn.Linear layers. No bias on Q, K, V (standard practice):"),
  codeBlock([
    "self.W_Q = nn.Linear(d_model, d_model, bias=False)  # speech → query space",
    "self.W_K = nn.Linear(d_model, d_model, bias=False)  # phoneme → key space",
    "self.W_V = nn.Linear(d_model, d_model, bias=False)  # phoneme → value space",
    "self.W_O = nn.Linear(d_model, d_model, bias=False)  # merge heads → output"
  ]),
  blank(),

  noteBox(
    "Write It Yourself — Exercise 2",
    "Write PhonemeCrossAttention(nn.Module). __init__ should define W_Q, W_K, W_V, W_O, norm, drop. " +
    "forward(self, H, P) where H: [B,T,F',C] and P: [B,T,C] should return [B,T,F',C]. " +
    "Steps: (1) reshape H and P to [BFp,T,C], (2) compute Q=H@W_Q, K=P@W_K, V=P@W_V, " +
    "(3) split heads, (4) compute attention scores and Z, (5) merge heads, " +
    "(6) apply W_O, residual, LayerNorm, (7) reshape back to [B,T,F',C].",
    PURPLE_BG, "5B2380"
  ),

  new Paragraph({ children: [new PageBreak()] }),

  // ── CHAPTER 3: Boundary-Biased MHSA ──────────────────────────────────────
  h1("Chapter 3 — Boundary-Biased MHSA (BB-MHSA)"),
  blank(),
  noteBox(
    "What you will write",
    "Replace the standard MHSA in CMGAN's TIME conformer block. " +
    "Add a phoneme similarity bias to the attention logits so that frames from different " +
    "phonemes attend less to each other, reducing correction smearing across boundaries. " +
    "Operates on [BFp, T, C], returns [BFp, T, C].",
    BLUE_BG, ACCENT
  ),
  blank(),

  h2("3.1  The core idea: additive attention bias"),
  p("Standard attention logits:"),
  codeBlock(["E_{i,j} = (q_i · k_j^T) / sqrt(d_k)"]),
  p("Modified (BB-MHSA) attention logits:"),
  codeBlock(["E_{i,j} = (q_i · k_j^T) / sqrt(d_k)  +  λ · S_{i,j}"]),
  p("where S_{i,j} is the phoneme similarity between frames i and j. High S_{i,j} → same phoneme → encourage attention. Low S_{i,j} → different phonemes → discourage attention."),
  blank(),

  h2("3.2  Computing the phoneme similarity matrix"),
  p("S is derived from CTC posteriors — probability distributions over phonemes at each frame:"),
  codeBlock([
    "# Step 1: Project C → n_phonemes+1 (the +1 is the CTC blank token)",
    "self.ctc_proj = nn.Linear(d_model, n_phonemes + 1, bias=False)",
    "",
    "# In forward():",
    "P_ctc = F.softmax(self.ctc_proj(P), dim=-1)   # [B, T, n_phonemes+1]",
    "#  Each row is a probability distribution over phonemes.",
    "#  High value at index k = model thinks phoneme k is being produced.",
    "",
    "# Step 2: Dot product of posterior distributions → similarity",
    "S = torch.bmm(P_ctc, P_ctc.transpose(1, 2))  # [B, T, T]",
    "#  S[i, j] = Σ_k P(i,k) * P(j,k)",
    "#  If i and j have similar posteriors → same phoneme → S close to 1",
    "#  If they have different posteriors → different phonemes → S close to 0"
  ]),
  blank(),

  h2("3.3  Broadcasting S across heads and F'"),
  p("S has shape [B, T, T]. The attention logits have shape [BFp, n_heads, T, T]. You need to broadcast:"),
  codeBlock([
    "# S: [B, T, T]",
    "# Need: [BFp, 1, T, T] so it broadcasts across n_heads",
    "",
    "Fp = BFp // B",
    "S_broadcast = (S.unsqueeze(1)             # [B, 1, T, T]",
    "                .expand(-1, Fp, -1, -1)   # [B, Fp, T, T]",
    "                .reshape(BFp, T, T)       # [BFp, T, T]",
    "                .unsqueeze(1))            # [BFp, 1, T, T]  ← broadcast over heads"
  ]),
  blank(),

  h2("3.4  The learnable lambda — initialise to zero"),
  p("λ is a learnable scalar. Critical: initialise to 0. This means at epoch 0 the model is identical to standard MHSA. During training, λ learns its optimal value."),
  codeBlock([
    "self.lam = nn.Parameter(torch.zeros(1))   # scalar, initialised to 0",
    "",
    "# In forward(), after computing base logits:",
    "logits = (Q @ K.transpose(-2, -1)) / (d_k ** 0.5)  # [BFp, n_heads, T, T]",
    "bias   = self.lam * S_broadcast                    # [BFp, 1, T, T]",
    "attn   = F.softmax(logits + bias, dim=-1)"
  ]),
  blank(),

  h2("3.5  Where to insert BB-MHSA in CMGAN"),
  p("In CMGAN's conformer block (generator.py), the time conformer has this structure:"),
  codeBlock([
    "# ORIGINAL CMGAN (simplified):",
    "class ConformerBlock(nn.Module):",
    "    def __init__(self):",
    "        self.mhsa   = nn.MultiheadAttention(...)  # ← REPLACE THIS",
    "        self.conv   = ConvModule(...)              # ← stays (for now)",
    "",
    "# DS-CMGAN:",
    "class DSConformerBlock(nn.Module):",
    "    def __init__(self):",
    "        self.bb_mhsa = BoundaryBiasedMHSA(...)    # ← your new class",
    "        self.conv    = ConvModule(...)"
  ]),
  blank(),

  noteBox(
    "Write It Yourself — Exercise 3",
    "Write BoundaryBiasedMHSA(nn.Module). forward(self, H_t, P) where H_t: [BFp,T,C] and P: [B,T,C]. " +
    "Steps: (1) compute S from ctc_proj and P using bmm, (2) broadcast S to [BFp,1,T,T], " +
    "(3) compute standard Q,K,V from H_t, (4) split into heads, (5) compute logits + lambda*S as bias, " +
    "(6) softmax → Z → merge heads → W_O → LayerNorm + residual. Return [BFp,T,C].",
    PURPLE_BG, "5B2380"
  ),

  new Paragraph({ children: [new PageBreak()] }),

  // ── CHAPTER 4: V/UV Gated Convolution ────────────────────────────────────
  h1("Chapter 4 — V/UV Gated Convolution Module (VUV-GC)"),
  blank(),
  noteBox(
    "What you will write",
    "Replace the ConvModule inside both time and frequency conformer stages. " +
    "The GLU gate is conditioned on voicing probability v(t): " +
    "voiced frames (vowels) use one gate behavior, unvoiced (consonants) use another. " +
    "Input/output: [batch, seq_len, C].",
    BLUE_BG, ACCENT
  ),
  blank(),

  h2("4.1  CMGAN's original ConvModule — understand it first"),
  p("The ConvModule in CMGAN's conformer block has this structure:"),
  codeBlock([
    "LayerNorm",
    "  ↓",
    "Pointwise Conv1d (1×1) — expands C → 2C",
    "  ↓",
    "GLU split: split 2C into x1 and x2, each C",
    "  gate = sigmoid(x2)   ← controls which features pass through",
    "  output = x1 * gate   ← gated output",
    "  ↓",
    "Depthwise Conv1d (kernel=31) — local feature capture",
    "  ↓",
    "Swish activation",
    "  ↓",
    "Pointwise Conv1d — projects back to C",
    "  ↓",
    "Dropout",
    "  ↓",
    "Residual connection: output + input"
  ]),
  blank(),

  h2("4.2  PyTorch Conv1d — the essential syntax"),
  codeBlock([
    "# Pointwise convolution (kernel_size=1) — like a per-position linear layer",
    "self.pw = nn.Conv1d(in_channels=64, out_channels=128, kernel_size=1)",
    "",
    "# Depthwise convolution — each channel processed independently",
    "self.dw = nn.Conv1d(",
    "    in_channels=64,",
    "    out_channels=64,",
    "    kernel_size=31,",
    "    padding=15,       # padding = kernel_size // 2  to keep length",
    "    groups=64         # groups=in_channels means depthwise",
    ")",
    "",
    "# IMPORTANT: Conv1d expects [batch, channels, length]",
    "# But our features are [batch, length, channels]",
    "# So always transpose before Conv1d and back after:",
    "x = x.transpose(1, 2)      # [B, T, C] → [B, C, T]",
    "x = self.conv(x)            # [B, C, T]",
    "x = x.transpose(1, 2)      # [B, C, T] → [B, T, C]",
    "",
    "# Chunking (for GLU split):",
    "x1, x2 = x.chunk(2, dim=1) # splits C=128 into two tensors of C=64"
  ]),
  blank(),

  h2("4.3  The voicing conditioning — the new part"),
  p("Original GLU gate:  gate = sigmoid(x2)"),
  p("VUV-GC gate:        gate = sigmoid(x2 + v_emb(t))"),
  blank(),
  p("v_emb maps [v(t), 1-v(t)] to C-dimensional space:"),
  codeBlock([
    "self.v_embed = nn.Linear(2, d_model)   # 2 inputs, C outputs",
    "",
    "# In forward():",
    "# v_prob: [B, seq_len]  ← voicing probability per frame",
    "v_in  = torch.stack([v_prob, 1.0 - v_prob], dim=-1)  # [B, seq_len, 2]",
    "v_emb = self.v_embed(v_in)                           # [B, seq_len, C]",
    "",
    "# Condition the gate:",
    "gate = torch.sigmoid(x2 + v_emb)      # [B, seq_len, C]",
    "# When v(t) → 1 (voiced): gate opens in 'vowel correction' mode",
    "# When v(t) → 0 (unvoiced): gate opens in 'consonant correction' mode"
  ]),
  blank(),

  h2("4.4  Getting voicing probability v(t)"),
  p("Two options:"),
  bullet("Option A — offline: compute v(t) using torchcrepe (pitch tracker) before training, save as .npy alongside the audio, load in the dataset class."),
  bullet("Option B — learned: a small 3-layer CNN trained jointly that predicts v(t) from the magnitude spectrogram. More end-to-end, fewer pre-processing steps."),
  blank(),
  p("Learnable V/UV classifier:"),
  codeBlock([
    "class LearnableVUV(nn.Module):",
    "    def __init__(self, n_freqs=200):",
    "        super().__init__()",
    "        self.net = nn.Sequential(",
    "            nn.Conv1d(n_freqs, 32, kernel_size=5, padding=2),",
    "            nn.ReLU(),",
    "            nn.Conv1d(32, 16, kernel_size=3, padding=1),",
    "            nn.ReLU(),",
    "            nn.Conv1d(16, 1,  kernel_size=1),",
    "            nn.Sigmoid()       # output ∈ [0, 1]",
    "        )",
    "",
    "    def forward(self, mag):    # mag: [B, T, F]",
    "        x = mag.transpose(1, 2)              # [B, F, T]",
    "        return self.net(x).squeeze(1)        # [B, T]"
  ]),
  blank(),

  h2("4.5  V/UV for the frequency conformer"),
  p("The frequency conformer operates on [BT, F', C] — the time dimension is folded into batch. Voicing is a time-domain property, so you pass a scalar per frame, broadcast over frequency bins:"),
  codeBlock([
    "# v_prob: [B, T]  ← one value per time frame",
    "",
    "# For frequency conformer: reshape B*T into batch",
    "v_avg  = v_prob.reshape(B * T, 1)          # [BT, 1]",
    "v_freq = v_avg.expand(-1, F_prime)          # [BT, F']",
    "# Each frequency bin at the same time frame gets the same voicing prob"
  ]),
  blank(),

  noteBox(
    "Write It Yourself — Exercise 4",
    "Write VUVGatedConvModule(nn.Module). __init__ should define: norm (LayerNorm), " +
    "pw1 (Conv1d: C→2C), dw (Conv1d depthwise: C→C, kernel=31), pw2 (Conv1d: C→C), " +
    "swish (nn.SiLU()), drop (Dropout), v_embed (Linear 2→C). " +
    "forward(self, x, v_prob) where x:[B,seq,C], v_prob:[B,seq]. " +
    "Key steps: save residual, LayerNorm, pointwise → GLU split, add v_emb to gate, " +
    "depthwise conv, Swish, pointwise, dropout, add residual. Return [B,seq,C].",
    PURPLE_BG, "5B2380"
  ),

  new Paragraph({ children: [new PageBreak()] }),

  // ── CHAPTER 5: CTC Loss ───────────────────────────────────────────────────
  h1("Chapter 5 — CTC Multi-Task Phoneme Head"),
  blank(),
  noteBox(
    "What you will write",
    "A lightweight head attached to the final TS-Conformer block output during training. " +
    "Forces the conformer to produce phoneme-discriminative representations. " +
    "Discarded at inference. Input: [B, T, F', C]. Output: scalar loss.",
    BLUE_BG, ACCENT
  ),
  blank(),

  h2("5.1  What CTC loss does"),
  p("CTC (Connectionist Temporal Classification) computes the probability of a phoneme label sequence given a sequence of per-frame distributions. It handles the alignment between frame-level predictions and phoneme-level labels automatically — you don't need to know which frames correspond to which phonemes."),
  blank(),
  p("PyTorch's CTCLoss:"),
  codeBlock([
    "loss_fn = nn.CTCLoss(",
    "    blank=39,             # index of the blank token (n_phonemes for 0-indexed)",
    "    reduction='mean',     # average over batch",
    "    zero_infinity=True    # replace inf losses with 0 (prevents NaN from misaligned seqs)",
    ")",
    "",
    "# Inputs:",
    "# log_probs:     [T, B, n_phonemes+1]   ← log-softmax of model output",
    "# targets:       [total_labels]          ← concatenated phoneme label sequences",
    "# input_lengths: [B]                     ← valid frames per utterance",
    "# target_lengths:[B]                     ← number of phonemes per utterance",
    "",
    "loss = loss_fn(log_probs, targets, input_lengths, target_lengths)"
  ]),
  blank(),

  h2("5.2  Preparing the input"),
  codeBlock([
    "# D_N: [B, T, F', C]  ← final conformer block output",
    "",
    "# Step 1: Average over frequency dimension → [B, T, C]",
    "D_avg = D_N.mean(dim=2)            # mean over F' = dim 2",
    "",
    "# Step 2: Project C → n_phonemes+1",
    "# self.proj = nn.Linear(C, n_phonemes + 1)  ← defined in __init__",
    "logits = self.proj(D_avg)          # [B, T, n_phonemes+1]",
    "",
    "# Step 3: Log-softmax and permute",
    "log_p  = F.log_softmax(logits, dim=-1)   # [B, T, n_phonemes+1]",
    "log_p  = log_p.permute(1, 0, 2)          # [T, B, n_phonemes+1]  ← CTCLoss format"
  ]),
  blank(),

  h2("5.3  Getting phoneme labels — Montreal Forced Aligner"),
  p("MFA aligns audio to transcription at the phoneme level. It produces .TextGrid files giving you the phoneme at each time interval. Convert these to integer sequences using ARPABET (39 phonemes)."),
  codeBlock([
    "# After running MFA on TORGO:",
    "# For each utterance you get a sequence like:",
    "# ['SIL', 'HH', 'EH', 'L', 'OW', 'SIL']  (for 'hello')",
    "# Map to integers:",
    "ARPABET = {'AA':0,'AE':1,'AH':2,'AO':3,'AW':4,'AY':5,'B':6,'CH':7,",
    "           'D':8,'DH':9,'EH':10,'ER':11,'EY':12,'F':13,'G':14,",
    "           'HH':15,'IH':16,'IY':17,'JH':18,'K':19,'L':20,'M':21,",
    "           'N':22,'NG':23,'OW':24,'OY':25,'P':26,'R':27,'S':28,",
    "           'SH':29,'T':30,'TH':31,'UH':32,'UW':33,'V':34,'W':35,",
    "           'Y':36,'Z':37,'ZH':38}  # 39 phonemes, blank=39",
    "",
    "# Pack sequences for CTCLoss:",
    "# targets = torch.cat([torch.tensor(seq) for seq in batch_seqs])  # [total_labels]",
    "# input_lengths = [T] * B  (number of frames)",
    "# target_lengths = [len(seq) for seq in batch_seqs]"
  ]),
  blank(),

  h2("5.4  Training-only activation"),
  codeBlock([
    "# In the generator's forward():",
    "ctc_loss = None",
    "if self.training and phoneme_seqs is not None:   # self.training is True during train()",
    "    ctc_loss = self.ctc_head(D_N, phoneme_seqs, input_lengths, label_lengths)",
    "",
    "# At inference: generator.eval()  →  self.training = False  →  CTC skipped"
  ]),
  blank(),

  noteBox(
    "Write It Yourself — Exercise 5",
    "Write CTCPhonemeHead(nn.Module). __init__ takes d_model and n_phonemes, " +
    "defines self.proj (Linear) and self.ctc (nn.CTCLoss with blank=n_phonemes, zero_infinity=True). " +
    "forward(self, D_N, phoneme_seqs, input_lengths, label_lengths): " +
    "(1) mean over dim=2 of D_N, (2) proj, (3) log_softmax + permute to [T,B,V], " +
    "(4) call self.ctc and return the loss.",
    PURPLE_BG, "5B2380"
  ),

  new Paragraph({ children: [new PageBreak()] }),

  // ── CHAPTER 6: Formant-Weighted Loss ─────────────────────────────────────
  h1("Chapter 6 — Formant-Weighted Magnitude Loss"),
  blank(),
  noteBox(
    "What you will write",
    "Replace CMGAN's uniform ||Xm - Xm_hat||² with a version that applies higher " +
    "weight at F1, F2, F3 frequency bins. F2 (~1560 Hz) gets the highest weight — " +
    "it's the most diagnostically important for vowel space compression in dysarthria.",
    BLUE_BG, ACCENT
  ),
  blank(),

  h2("6.1  CMGAN's original magnitude loss"),
  codeBlock([
    "# CMGAN (from losses.py):",
    "L_mag = F.mse_loss(Xm, Xm_hat)   # uniform weight at all frequencies",
    "",
    "# This treats a 1dB error at 100 Hz the same as a 1dB error at 1560 Hz.",
    "# For dysarthric speech, errors at F2 are more perceptually damaging."
  ]),
  blank(),

  h2("6.2  The Gaussian frequency weight vector"),
  p("For CMGAN's STFT: 16kHz audio, 400-point FFT → frequency resolution = 16000/400 = 40 Hz per bin. F bins = 201 (400//2 + 1)."),
  codeBlock([
    "# Bin numbers for formant centres:",
    "# F1 ≈  530 Hz  →  530 / 40 = bin 13",
    "# F2 ≈ 1560 Hz  → 1560 / 40 = bin 39   ← highest weight",
    "# F3 ≈ 2840 Hz  → 2840 / 40 = bin 71",
    "# HF ≈ 5600 Hz  → 5600 / 40 = bin 140  ← fricatives",
    "",
    "# Gaussian bump:  w_i * exp( -(k - mu_i)^2 / (2 * sigma_i^2) )",
    "import torch",
    "F_bins = 201",
    "k = torch.arange(F_bins, dtype=torch.float32)   # [0, 1, 2, ..., 200]",
    "",
    "W = torch.ones(F_bins)   # start at 1.0 everywhere",
    "formants = [(13, 6, 0.8), (39, 10, 1.0), (71, 8, 0.6), (140, 30, 0.4)]",
    "for mu, sigma, w in formants:",
    "    W += w * torch.exp(-(k - mu)**2 / (2 * sigma**2))",
    "",
    "W = W / W.mean()   # normalise so total loss scale matches CMGAN"
  ]),
  blank(),

  h2("6.3  Applying the weight in the loss"),
  codeBlock([
    "# W: [F_bins]   — register as buffer (not parameter)",
    "# Xm_clean, Xm_hat: [B, T, F]",
    "",
    "diff = (Xm_clean - Xm_hat) ** 2    # [B, T, F]",
    "",
    "# W needs to broadcast over B and T:",
    "# diff * self.W  works automatically because PyTorch broadcasts",
    "# from the right: [B, T, F] * [F] → element-wise on the F dimension",
    "",
    "loss = (diff * self.W).mean()       # scalar"
  ]),
  blank(),

  noteBox(
    "Write It Yourself — Exercise 6",
    "Write FormantWeightedLoss(nn.Module). __init__ builds W using the formant list above " +
    "and registers it as a buffer (not a parameter). " +
    "forward(self, Xm_clean, Xm_hat) computes and returns the weighted MSE. " +
    "Verify: (a) W.shape == (201,). (b) W.mean() ≈ 1.0. (c) 'W' not in dict(model.named_parameters()).",
    PURPLE_BG, "5B2380"
  ),

  new Paragraph({ children: [new PageBreak()] }),

  // ── CHAPTER 7: Putting it together ───────────────────────────────────────
  h1("Chapter 7 — Putting It Together: The DS-Conformer Block"),
  blank(),
  p("Now you have all the pieces. The DSConformerBlock replaces CMGAN's TSConformer by combining your four modules inside the existing time→frequency structure."),
  blank(),

  h2("7.1  CMGAN's original TSConformer — read this first"),
  p("In CMGAN's generator.py, the TSConformer works as follows:"),
  codeBlock([
    "class TSConformer(nn.Module):",
    "    def __init__(self):",
    "        # Time conformer",
    "        self.time_conformer = ConformerBlock()",
    "        # Frequency conformer",
    "        self.freq_conformer = ConformerBlock()",
    "",
    "    def forward(self, x):                      # x: [B, T, F', C]",
    "        B, T, Fp, C = x.shape",
    "",
    "        # TIME: fold F' into batch",
    "        xt = x.permute(0,2,1,3).reshape(B*Fp, T, C)",
    "        xt = self.time_conformer(xt)",
    "        xt = xt.reshape(B, Fp, T, C).permute(0,2,1,3)  # back to [B,T,F',C]",
    "        x  = x + xt                               # residual",
    "",
    "        # FREQUENCY: fold T into batch",
    "        xf = x.reshape(B*T, Fp, C)",
    "        xf = self.freq_conformer(xf)",
    "        xf = xf.reshape(B, T, Fp, C)",
    "        x  = x + xf                               # residual",
    "",
    "        return x"
  ]),
  blank(),

  h2("7.2  What changes for DS-CMGAN"),
  codeBlock([
    "class DSConformerBlock(nn.Module):",
    "    def __init__(self):",
    "        # Time conformer — MHSA replaced with BB-MHSA",
    "        self.time_bb_mhsa = BoundaryBiasedMHSA()",
    "        self.time_vuv_gc  = VUVGatedConvModule()  # replaces ConvModule",
    "        self.time_ffn1    = ...   # same as CMGAN",
    "        self.time_ffn2    = ...   # same as CMGAN",
    "",
    "        # Frequency conformer — standard MHSA, but VUV-GC replaces ConvModule",
    "        self.freq_mhsa    = nn.MultiheadAttention(...)",
    "        self.freq_vuv_gc  = VUVGatedConvModule()",
    "        self.freq_ffn1    = ...",
    "        self.freq_ffn2    = ...",
    "",
    "        # PP-CA — after the full block",
    "        self.pp_ca = PhonemeCrossAttention()",
    "",
    "    def forward(self, x, P, v_prob):           # two new arguments: P and v_prob",
    "        # ... time conformer with BB-MHSA and VUV-GC ...",
    "        # ... freq conformer with standard MHSA and VUV-GC ...",
    "        # ... PP-CA at the end ...",
    "        return x"
  ]),
  blank(),

  h2("7.3  How to pass P and v_prob through the stack"),
  p("In CMGAN's generator, N=4 TSConformer blocks are stacked in a loop. You need to pass P and v_prob to every block:"),
  codeBlock([
    "# CMGAN original (generator.py, inside Generator.forward()):",
    "for block in self.ts_conformers:",
    "    x = block(x)",
    "",
    "# DS-CMGAN:",
    "P      = self.phoneme_extractor(waveform, T)   # compute once, reuse in every block",
    "v_prob = self.vuv_classifier(Ym)               # compute once, reuse in every block",
    "",
    "for block in self.ds_blocks:",
    "    x = block(x, P, v_prob)                    # pass to every block"
  ]),
  blank(),

  h2("7.4  The complete loss function"),
  codeBlock([
    "# In losses.py / train.py:",
    "",
    "L_mag  = formant_loss(Xm_clean, Xm_hat)         # your FormantWeightedLoss",
    "L_ri   = F.mse_loss(Xr_clean, Xr_hat) + F.mse_loss(Xi_clean, Xi_hat)",
    "L_tf   = 0.7 * L_mag + 0.3 * L_ri               # same alpha as CMGAN",
    "",
    "L_time = F.l1_loss(x_clean, x_hat)              # same as CMGAN",
    "L_gan  = F.mse_loss(D(Xm, Xm_hat), ones)        # same as CMGAN",
    "",
    "L_ctc  = ctc_loss if ctc_loss is not None else 0.0",
    "",
    "# Warmup schedule: gamma4 = 0 for first 5 epochs, ramps to 0.3 by epoch 20",
    "L_G = gamma1*L_tf + gamma2*L_gan + gamma3*L_time + gamma4*L_ctc"
  ]),
  blank(),

  h2("7.5  Warmup schedule for CTC weight"),
  codeBlock([
    "def get_gamma4(epoch):",
    "    if epoch < 5:",
    "        return 0.0          # CTC off — let base losses stabilise first",
    "    elif epoch < 20:",
    "        t = (epoch - 5) / 15.0",
    "        return 0.3 * t      # linear ramp: 0.0 → 0.3",
    "    else:",
    "        return 0.3          # full weight",
    "",
    "# In train loop:",
    "gamma4 = get_gamma4(current_epoch)"
  ]),
  blank(),

  new Paragraph({ children: [new PageBreak()] }),

  // ── CHAPTER 8: Testing Your Code ─────────────────────────────────────────
  h1("Chapter 8 — How to Test Every Module You Write"),
  blank(),
  p("For every module you write, run these four checks before moving to the next one."),
  blank(),

  h2("8.1  Shape check — the minimum bar"),
  codeBlock([
    "# After writing any module MyModule:",
    "m = MyModule().to('cuda')",
    "dummy = torch.randn(4, 320, 100, 64).to('cuda')   # B=4, T=320, F'=100, C=64",
    "",
    "out = m(dummy, ...)",
    "assert out.shape == dummy.shape, f'Expected {dummy.shape}, got {out.shape}'",
    "print('Shape check passed')"
  ]),
  blank(),

  h2("8.2  Gradient flow check — verify backprop works"),
  codeBlock([
    "# Run a fake backward pass",
    "loss = out.sum()",
    "loss.backward()",
    "",
    "# Every trainable parameter should have a gradient",
    "for name, param in m.named_parameters():",
    "    if param.requires_grad:",
    "        assert param.grad is not None, f'No gradient for {name}'",
    "        print(f'{name}: gradient OK  (norm={param.grad.norm().item():.4f})')",
    "",
    "# For FrozenPhoneticExtractor: verify Wav2Vec2 has NO gradient",
    "for p in extractor.wav2vec2.parameters():",
    "    assert p.grad is None, 'Wav2Vec2 should not have gradients'"
  ]),
  blank(),

  h2("8.3  Finiteness check — no NaN or Inf"),
  codeBlock([
    "assert torch.isfinite(out).all(), 'Output contains NaN or Inf'",
    "assert torch.isfinite(loss), f'Loss is not finite: {loss.item()}'"
  ]),
  blank(),

  h2("8.4  Lambda initialisation check — for BB-MHSA"),
  codeBlock([
    "bb = BoundaryBiasedMHSA()",
    "assert bb.lam.item() == 0.0, f'Lambda should be 0 at init, got {bb.lam.item()}'",
    "print(f'Lambda at init: {bb.lam.item()} (correct: 0.0)')"
  ]),
  blank(),

  h2("8.5  Buffer vs parameter check — for FormantWeightedLoss"),
  codeBlock([
    "loss_fn = FormantWeightedLoss()",
    "",
    "# W should be a buffer, not a parameter",
    "param_names = [n for n, _ in loss_fn.named_parameters()]",
    "buffer_names = [n for n, _ in loss_fn.named_buffers()]",
    "",
    "assert 'W' not in param_names, 'W should NOT be a parameter'",
    "assert 'W'     in buffer_names, 'W SHOULD be a buffer'",
    "print('Buffer check passed')",
    "",
    "# W should be on GPU after .to(device)",
    "loss_fn = loss_fn.to('cuda')",
    "assert loss_fn.W.device.type == 'cuda', 'W not on GPU'"
  ]),
  blank(),

  new Paragraph({ children: [new PageBreak()] }),

  // ── CHAPTER 9: Integration with CMGAN ────────────────────────────────────
  h1("Chapter 9 — Integration Map: Where Each Module Goes in CMGAN"),
  blank(),
  p("This chapter tells you exactly which file to open and which line to change for each module."),
  blank(),

  new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [2800, 2200, 4360],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders, shading: { fill: ACCENT, type: ShadingType.CLEAR },
          width: { size: 2800, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: "Module", font: FONT, size: 20, bold: true, color: "FFFFFF" })] })] }),
        new TableCell({ borders, shading: { fill: ACCENT, type: ShadingType.CLEAR },
          width: { size: 2200, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: "File", font: FONT, size: 20, bold: true, color: "FFFFFF" })] })] }),
        new TableCell({ borders, shading: { fill: ACCENT, type: ShadingType.CLEAR },
          width: { size: 4360, type: WidthType.DXA },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: "What to change", font: FONT, size: 20, bold: true, color: "FFFFFF" })] })] }),
      ]}),
      ...[
        ["FrozenPhoneticExtractor", "models/generator.py", "Add as self.phoneme_extractor in Generator.__init__. Call in forward() before conformer stack."],
        ["LearnableVUV", "models/generator.py", "Add as self.vuv_classifier in Generator.__init__. Call in forward() on Ym."],
        ["BoundaryBiasedMHSA", "models/conformer.py", "In ConformerBlock.__init__: replace self.self_attn = nn.MultiheadAttention with self.bb_mhsa = BoundaryBiasedMHSA. Update forward() to pass P."],
        ["VUVGatedConvModule", "models/conformer.py", "In ConformerBlock.__init__: replace self.conv_module = ConvModule with self.vuv_gc = VUVGatedConvModule. Update forward() to pass v_prob."],
        ["PhonemeCrossAttention", "models/conformer.py", "Add self.pp_ca = PhonemeCrossAttention() to TSConformer. Call at end of TSConformer.forward()."],
        ["CTCPhonemeHead", "models/generator.py", "Add self.ctc_head = CTCPhonemeHead() to Generator. Call on D_N (final conformer output) during training."],
        ["FormantWeightedLoss", "losses.py", "Replace F.mse_loss(Xm, Xm_hat) in the magnitude loss with FormantWeightedLoss.forward(Xm, Xm_hat)."],
      ].map((row, i) => new TableRow({ children: row.map((cell, j) => new TableCell({
        borders,
        shading: { fill: i % 2 === 0 ? GRAY_BG : "FFFFFF", type: ShadingType.CLEAR },
        width: { size: [2800,2200,4360][j], type: WidthType.DXA },
        margins: { top: 60, bottom: 60, left: 120, right: 120 },
        children: [new Paragraph({ spacing:{before:0,after:0}, children: [new TextRun({ text: cell, font: j===1?CODE_FONT:FONT, size: j===1?18:20 })] })]
      })) }))
    ]
  }),

  blank(), blank(),

  h2("9.1  The order to write and test everything"),
  bullet("Week 1: Write FrozenPhoneticExtractor. Run shape + gradient checks. Show to Prof. Patil."),
  bullet("Week 1: Write FormantWeightedLoss. Run buffer and loss checks."),
  bullet("Week 2: Write VUVGatedConvModule and LearnableVUV. Integrate into ConformerBlock."),
  bullet("Week 2: Write BoundaryBiasedMHSA. Integrate into time conformer. Check λ init."),
  bullet("Week 3: Write PhonemeCrossAttention. Insert after each TSConformer block."),
  bullet("Week 3: Write CTCPhonemeHead. Integrate into Generator. Implement warmup schedule."),
  bullet("Week 4: Run full forward + backward pass on dummy data. All gradient checks."),
  bullet("Week 4: Retrain CMGAN baseline on TORGO. Start DS-CMGAN ablation training."),
  blank(),

  noteBox(
    "Final advice",
    "Write each module in a separate .py file first. Test it fully with dummy tensors before " +
    "integrating it into CMGAN. When you integrate, change the minimum possible in CMGAN's " +
    "original code — add new arguments (P, v_prob) and new calls, but don't restructure " +
    "CMGAN's existing logic. This keeps your diffs small and makes debugging much easier.",
    GREEN_BG, GREEN
  ),

];

// ── Build and save document ───────────────────────────────────────────────────
const doc = new Document({
  numbering: {
    config: [{
      reference: "bullets",
      levels: [{
        level: 0,
        format: LevelFormat.BULLET,
        text: "•",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 720, hanging: 360 } } }
      }, {
        level: 1,
        format: LevelFormat.BULLET,
        text: "◦",
        alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 1080, hanging: 360 } } }
      }]
    }]
  },
  styles: {
    default: { document: { run: { font: FONT, size: 22 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 32, bold: true, font: FONT, color: ACCENT },
        paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 26, bold: true, font: FONT, color: ACCENT2 },
        paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: FONT, color: "333333" },
        paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 2 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 12240, height: 15840 },
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    children
  }]
});

Packer.toBuffer(doc).then(buffer => {
  fs.writeFileSync('./DS_CMGAN_Teaching_Guide.docx', buffer);
  console.log('Done: DS_CMGAN_Teaching_Guide.docx');
});