**System Identification using Excel + macros**

For a system expressed in its Laplace representation:

<img src="https://latex.codecogs.com/svg.latex?\Large&space;x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}" title="\Large x=\frac{-b\pm\sqrt{b^2-4ac}}{2a}" />

We want to identify the parameters $a_1$, $a_2$ and $a_3$ . Bilineal transformation is used to obtain a quasi-equivalent Zeta-transform which difference equation is:





In the real system a voltage $V_{exp}$ (t) is applied to obtain the angular output $\theta_{exp}(t)$ .

We use the solver addin of MS. Excel to minimize the cost function:





At the minimum $J$ the corresponding values of $a*$, $b*$ and $c*$ are the optimal values we were looking for.
