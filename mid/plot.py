import numpy as np
import matplotlib.pyplot as plt
# phi = (1 + 5 ** 0.5) / 2

# t = np.linspace(0,100*np.pi,10000)
# x = np.cos(t)
# y = -np.sin(np.pi*t)

# plt.plot(x,y,color='k')
# plt.axis('equal')
# plt.show()

rb = 1                                  # base radius
psi = np.linspace(0,np.pi, 100)         # roll angle
x = rb*(np.cos(psi)+psi*np.sin(psi))    # cartesian x-coordinate
y = rb*(np.sin(psi)-psi*np.cos(psi))    # cartesian y-coordinate
R = 1                                   # radius of curvature
k = 1/R                                   # curvature
alpha = 1                               # pressure angle
invalpha = 1                            # involute function of pressure angle alpha

psic = np.linspace(0,2*np.pi, 100)
xc = rb*np.cos(psic)
yc = rb*np.sin(psic)

angles = np.linspace(0,1,100)*np.pi

for psi in angles:
    plt.plot([rb*np.cos(psi),rb*(np.cos(psi)+psi*np.sin(psi))],[rb*np.sin(psi),rb*(np.sin(psi)-psi*np.cos(psi))],'r')
plt.plot(x,y, color='k')
plt.plot(xc, yc,color='b')



rb = 1                                  # base radius
psi = np.linspace(0,np.pi, 100)         # roll angle
x = 2*rb-rb*(np.cos(psi)+psi*np.sin(psi))    # cartesian x-coordinate
y = -rb*(np.sin(psi)-psi*np.cos(psi))    # cartesian y-coordinate
R = 1                                   # radius of curvature
k = 1/R                                   # curvature
alpha = 1                               # pressure angle
invalpha = 1                            # involute function of pressure angle alpha

psic = np.linspace(0,2*np.pi, 100)
xc = 2*rb-rb*np.cos(psic)
yc = -rb*np.sin(psic)

angles = np.linspace(0,1,100)*np.pi

for psi in angles:
    plt.plot([2*rb-rb*np.cos(psi),2*rb-rb*(np.cos(psi)+psi*np.sin(psi))],[-rb*np.sin(psi),-rb*(np.sin(psi)-psi*np.cos(psi))],'r')
plt.plot(x,y, color='k')
plt.plot(xc, yc,color='b')
plt.axis('equal')
plt.grid()




plt.show()

