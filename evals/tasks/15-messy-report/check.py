from evalkit import Checks, args, mentions

work, truth, answer = args()
c = Checks()
c.add("West Q2 sales in dollars, without the discontinued product", mentions(answer, truth["west_q2"], rel=0.001), f"want {truth['west_q2']:,}")
# Bravo has the largest single row; Alpha wins if the figure stored as text
# is left out.
c.add(f"the top Q2 product is {truth['best']}", truth["best"] in answer, f"q2 {truth['q2']}")
c.report()
