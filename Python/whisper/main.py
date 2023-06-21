import whisper
import numpy as np
import torch

torch.cuda.is_available()
DEVICE = "cuda" if torch.cuda.is_available() else "cpu"

model = whisper.load_model("tiny.en", device=DEVICE)

print(
    f"Model is {'multilingual' if model.is_multilingual else 'English-only'} "
    f"and has {sum(np.prod(p.shape) for p in model.parameters()):,} parameters."
)

result = model.transcribe("c:\\Users\\majona\\Downloads\\LL_1st_alert.aac")

print(result["text"])

with open('readme.txt', 'w') as f:
    f.writelines(result["text"])