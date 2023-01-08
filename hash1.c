// Hashing
// Uses special Hash function
// A hash function maps element to an address for storage
// This provides constant-time access
// Collision is handled by collision resolution techniques
// Collision resolution technique
// Chaining
// Open Addressing
// ----------------------------------------------------------Advantages
// The hash function helps in fetching element in constant time
// An efficient way to store elements
// ----------------------------------------------------------Disadvantages
// Collision resolution increases complexity
// Applications
// Suitable for the application needs constant time fetching
// Graph
// Basically it is a group of edges and vertices
// Graph representation
// G(V, E): where V(G) represents a set of vertices and E(G) represents a set of edges
// A graph can be directed or undirected
// A graph can be connected or disjoint
// -----------------------------------------------------------Advantages
// finding connectivity
// Shortest path
// min cost to reach from 1 pt to other
// Min spanning tree
// ----------------------------------------------------------Disadvantages
// Storing graph(Adjacency list and Adjacency matrix) can lead to complexities
// Applications
// Suitable for a circuit network
// Suitable for applications like Facebook, LinkedIn, etc
// Medical science

#include <stdio.h>
#include <stdlib.h>

struct node
{
    int vertex;
    struct node *next;
};
struct node *createNode(int);

struct Graph
{
    int numVertices;
    struct node **adjacentLists;
};

struct node *createNode(int v)
{
    struct node *newNode = malloc(sizeof(struct node));
    newNode->vertex = v;
    newNode->next = NULL;
    return newNode;
}

struct Graph *createAGraph(int vertices)
{
    struct Graph *graph = malloc(sizeof(struct Graph));
    graph->numVertices = vertices;

    graph->adjacentLists = malloc(vertices * sizeof(struct node *));

    int i;
    for (i = 0; i < vertices; i++)
        graph->adjacentLists[i] = NULL;

    return graph;
}

void addEdge(struct Graph *graph, int a, int b)
{
    struct node *newNode = createNode(b);
    newNode->next = graph->adjacentLists[a];
    graph->adjacentLists[a] = newNode;

    newNode = createNode(a);
    newNode->next = graph->adjacentLists[b];
    graph->adjacentLists[b] = newNode;
}

void printGraph(struct Graph *graph)
{
    int v;
    for (v = 0; v < graph->numVertices; v++)
    {
        struct node *temp = graph->adjacentLists[v];
        printf("\n Vertex %d\n: ", v);
        while (temp)
        {
            printf("%d -> ", temp->vertex);
            temp = temp->next;
        }
        printf("\n");
    }
}

int main()
{
    struct Graph *graph = createAGraph(4);
    addEdge(graph, 0, 1);
    addEdge(graph, 0, 3);
    addEdge(graph, 0, 2);
    addEdge(graph, 1, 3);

    printGraph(graph);

    return 0;
}