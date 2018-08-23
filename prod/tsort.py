#    topological sort orders a directed acyclic graph,
#    in order to make sure nodes are only visited after
#    all their required nodes have been visited.

class Node:

    def __init__(self, name):
        self.name = name
        self.visited = False
        self.dependents = []

    def id(self):
        return self.name.__hash__()

class CycleFoundException(Exception):
    def __init__(self, data):
        self.data = repr(data)
		
def visit( n, L, S, V):
    if n.name in V:
        raise CycleFoundException((n.name, L))
    V.append( n.name )
    if not n.visited:
        n.visited = True
        for d in n.dependents:
            visit( S[d], L, S, V )
        L.append( n.name )
    V.remove( n.name )

#
# Perform a topolgical sort.
# arcs is a collection of pairs (a, b) where b depends on a
# or in other words, a must be processed before b
# Returns a list sorted topologicaly
# or throws a fault if there is a cycle.
#
def tsort(arcs):
    L = [] # sorted list
    S = {} # Set of all nodes
    for a,b in arcs:
        if not a in S:
            S[a] = Node( a )
        if not b in S:
            S[b] = Node( b )
        if a not in S[b].dependents:
            S[b].dependents.append( a ) # b depends on a
		
    try:
	    for n in S.itervalues():
		V = [] # list of Visited nodes used to detect loops
		visit( n, L, S, V )
    except CycleFoundException as e:
	    print 'ERROR: Cycle found at: {}'.format( e.data )
	    return None

    return L
	
#--------------------------------------------------------------

def test():    
    list = [ ('a','b'), ('c','d'), ('a','f'), ('f','c'), ('c','i'), ('d','b') ]
    
    # a correct order is ['a', 'f', 'c', 'd', 'b', 'i']
    print list
    print "Expecting ['a', 'f', 'c', 'd', 'b', 'i']"
    print "Obtained: {}".format( tsort(list) )
    print
        
    list = [ ('a','b'), ('c','d'), ('a','f'), ('f','c'), ('b','c'), ('d','b') ]
    # cycle : b c, c d, d b
    print "Expecting a cycle at ('c', ['a', 'f'], ['c', 'b', 'd']):"
    print "Obtained: {}".format( tsort(list) )

if __name__ == "__main__":
    test()       
