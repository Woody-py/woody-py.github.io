---
layout: post
title:  "How to derive Spot Rate from a Swap Curve with Excel/VBA"
date:   2020-08-31 00:45:00 +0900
categories: jekyll update
---
Hi all,

Deriving spot rate from a swap curve is one of the most basic exercise for finance, using

basic principles such as *"No Arbitrage"*

Let's see how to do this with Excel/VBA.

These are contents for this post.

- Basic Idea : No Arbitrage
- Basic Concepts for Excel Solver.  GRG vs evolutionary (Optional)
- Deriving spot rates with Excel Solver
- Deriving spot rates with VBA and analytic solution



## Basic Idea : No Arbitrage

No arbitrage can be expressed in various ways. For this exercise, I would like to say that

 it means "the same prices mean the same expected cash flows."
$$
100 = \left(\frac{swapRate}{n} \times 100 \times \sum_{j=1}^{M} Z\left(0, T_{j}\right)+Z\left(0, T_{M}\right) \times 100\right)
$$
It can be expressed in this formula for this exercise. Z stands for zero rate which is our target. 

Left side means the value of floating rate coupon bond will be 100 on reset dates. This is only correct under the assumption  that floating rate and discount factor have inverse relation. It will be justified if that floating rate can be deemed as risk-free rate such as LIBOR, SOFR and SONIA. 

Right side just shows discounted coupons and the par value of the bond. 



## Basic Concepts for Excel Solver (Optional)

Now we are going to see basic concepts for Excel Solver, especially on GRG nonlinear solver and evolutionary solver. This part is optional since this is about algorithm or computer science rather than finance. If you are not interested these topics just remember you are going to use GRG nonlinear this time. 



### GRG nonlinear

GRG stands for "Generalized Reduced Gradient." You might be more familiar with gradient descent  if you are interested in machine learning. GRG nonlinear is almost same with gradient descent. 

![GRG Image](https://engineerexcel.com/wp-content/uploads/2016/11/112616_1408_ExcelSolver2.png)

( https://engineerexcel.com/wp-content/uploads/2016/11/112616_1408_ExcelSolver2.png)

As you can see in this picture, GRG solver change parameters to the direction of gradient (or slope) to minimize the difference with your target (usually out target is 0). It will stop if difference is not changing for some continuous times. It is fast but has the local minima problem. 



### Evolutionary

Evolutionary method(A.K.A. genetic algorithm) is stemmed from the way how genes evolve as you can see in its name. It is heuristic approach for combinatorial problems or I may say that any problems you failed to solve with GRG method.



Evolutionary algorithm contains four steps, 

1. Initialization
2. Selection
3. Genetic Operation
4. Termination

Let's see with a nice diagram from Max Maxfield's journal. (https://www.eejournal.com/article/when-genetic-algorithms-meet-artificial-intelligence/)

![Genetic Algorithm](https://www.eejournal.com/wp-content/uploads/2020/07/max-0040-02-genetic-algorithms-1024x608.png)

(https://www.eejournal.com/wp-content/uploads/2020/07/max-0040-02-genetic-algorithms-1024x608.png)

 As you can see, 

- first initialize genes(which are parameters we should optimize) on a random basis.  
- and then select preserved genes based on your loss or target function. 
- After that is breeding,  mix genes (breeding) and change small portion of genes randomly.(mutation). Mutation is for preventing bias from initial population.
- Keep this process until loss or target is not changing.



## Deriving spot rates with Excel Solver













## Deriving spot rates with VBA and analytic solution













Check out the [Jekyll docs][jekyll-docs] for more info on how to get the most out of Jekyll. File all bugs/feature requests at [Jekyll’s GitHub repo][jekyll-gh]. If you have questions, you can ask them on [Jekyll Talk][jekyll-talk].

[jekyll-docs]: https://jekyllrb.com/docs/home
[jekyll-gh]:   https://github.com/jekyll/jekyll
[jekyll-talk]: https://talk.jekyllrb.com/