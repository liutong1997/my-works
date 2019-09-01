import numpy as np
import matplotlib.pyplot as plt
# This import registers the 3D projection, but is otherwise unused.
from mpl_toolkits.mplot3d import Axes3D  # noqa: F401 unused import


# setup the figure and axes
fig = plt.figure(figsize=(6, 6))
ax1 = fig.add_subplot(111, projection='3d')


# fake data
_x = np.arange(2)
_y = np.arange(3)
_xx, _yy = np.meshgrid(_x, _y)
print(_xx)
x, y = _xx.ravel(), _yy.ravel()

top = np.array([1,2,3,4,5,6])
print(x)
print(y)
print(top)
bottom = np.zeros_like(top)
width = depth = 1

ax1.bar3d(x, y, bottom, width, depth, top, shade=True)
ax1.set_title('Shaded')


plt.show()