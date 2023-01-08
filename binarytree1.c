// Binary Tree
// Hierarchical  Data Structures using C
// Topmost element is known as the root of the tree
// Every node can have at most 2 children in the binary tree
// Can access elements randomly using index
// Eg: File system hierarchy
// Common traversal methods:
// preorder(root) : print-left-right
// postorder(root) : left-right-print
// inorder(root) : left-print-right
// ---------------------------------------------Advantages
// Can represent data with some relationship
// Insertion and search are much efficient
// ---------------------------------------------Disadvantages
// Sorting is difficult
// Not much flexible
// Applications
// File system hierarchy
// Multiple variations of the binary tree have a wide variety of applications

#include <stdio.h>
#include <conio.h>
#include <stdlib.h>

struct bst
{
    int data;
    struct bst *left;
    struct bst *right;
};

struct bst *insert(struct bst *, int);
void inorder(struct bst *);
void preorder(struct bst *);
void postorder(struct bst *);

int main()
{
    struct bst *r = NULL;
    r = insert(r, 30);
    r = insert(r, 15);
    r = insert(r, 10);
    r = insert(r, 20);
    r = insert(r, 40);
    r = insert(r, 5);
    r = insert(r, 45);
    r = insert(r, 35);
    printf("\n display element in inorder:-");
    inorder(r);
    printf("\n display element in preorder:-");
    preorder(r);
    printf("\n display element in postorder:-");
    postorder(r);
    return 1;
}

struct bst *insert(struct bst *q, int val)
{
    struct bst *tmp;
    tmp = (struct bst *)malloc(sizeof(struct bst));

    if (q == NULL)
    {
        tmp->data = val;
        tmp->left = tmp->right = NULL;
        return tmp;
    }
    else
    {
        if (val < (tmp->data))
        {
            q->left = insert(q->left, val);
        }
        else
        {
            q->right = insert(q->right, val);
        }
    }
    return q;
}

void inorder(struct bst *q)
{

    if (q == NULL)
    {
        return;
    }

    inorder(q->left);
    printf(" %d ", q->data);
    inorder(q->right);
}

void preorder(struct bst *q)
{

    if (q != NULL)
    {
        printf(" %d ", q->data);
        preorder(q->left);
        preorder(q->right);
    }
}

void postorder(struct bst *q)
{

    if (q != NULL)
    {
        postorder(q->left);
        postorder(q->right);
        printf(" %d ", q->data);
    }
}