**System Identification using Excel + macros**

For a system expressed in its Laplace representation:

$ \theta(s) =  \frac{1}{a_3 s^3+a_2 s^2+a_1 s}  V(s) $

We want to identify the parameters $a_1$, $a_2$ and $a_3$ . Bilineal transformation is used to obtain a quasi-equivalent Zeta-transform which difference equation is:

$\theta[n] = a*V[n-3] + b*\theta[n-3] + c*\theta[n-2] + d*\theta[n-1]$

In the real system a voltage $V_{exp}$ (t) is applied to obtain the angular output $\theta_{exp}(t)$ .

We use the solver addin of MS. Excel to minimize the cost function:

$J(a,b,c)=\sum{(\theta_{exp}[n] - \theta[n])^2} $

At the minimum $J$ the corresponding values of $a*$, $b*$ and $c*$ are the optimal values we were looking for.

![alt text](https://raw.githubusercontent.com/username/projectname/branch/path/to/img.png)
