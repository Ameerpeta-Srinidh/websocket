words = [line.split(',', 1)[1] for line in open('Indian_Names.csv').read().splitlines()[1:]]
print(len(words))
print(words[:10])


import torch

N = torch.zeros((28,28), dtype=torch.int32)
chars = sorted(list(set(''.join(words))))
stoi = {s:i for i,s in enumerate(chars)}
stoi['<E>'] = 27
stoi['<S>'] = 26