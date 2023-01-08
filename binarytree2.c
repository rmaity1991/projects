// Binary Search Tree
// A binary tree with the additional restriction
// ------------------------------------------Restriction:
// The left child must always be less than the root node
// The right child must always be greater than the root node
// Insertion, Deletion, Search is much more efficient than a binary tree
// ------------------------------------------Advantages
// Maintains order in elements
// Can easily find the min and max nodes in the tree
// In order traversal gives sorted elements
// ------------------------------------------Disadvantages
// Random access not possible
// Ordering adds complexity
// Applications
// Suitable for sorted hierarchical data

#include <stdio.h>
#include <stdlib.h>
struct bst
{
    int data;
    struct bst *left;
    struct bst *right;
};

struct bst *insert(struct bst *q, int val)
{
    struct bst *temp;
    if (q == NULL)
    {
        temp = (struct bst *)malloc(sizeof(struct bst));
        temp->data = val;
        temp->left = NULL;
        temp->right = NULL;
        q = temp;
    }
    else
    {
        if (val < q->data)
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
    printf("%d\t", q->data);
    inorder(q->right);
}

struct bst *search(struct bst *p, int key, struct bst **y)
{
    struct bst *temp;
    if (p == NULL)
        return (NULL);
    temp = p;
    *y = NULL;
    while (temp != NULL)
    {
        if (temp->data == key)
            return (temp);
        else
        {
            *y = temp; /*store this pointer as root */
            if (temp->data > key)
                temp = temp->left;
            else
                temp = temp->right;
        }
    }
    return (NULL);
}

/* A function to delete the node whose data value is given */
struct bst *del(struct bst *p, int val)
{
    struct bst *x, *y, *temp;
    x = search(p, val, &y);
    if (x == NULL)
    {
        printf("The node does not exists\n");
        return (p);
    }
    else
    {
        /* this code is for deleting root node*/
        if (x == p)
        {
            temp = x->left;
            y = x->right;
            p = temp;
            while (temp->right != NULL)
                temp = temp->right;
            temp->right = y;
            free(x);
            return (p);
        }
        /* this code is for deleting node having both children */
        if (x->left != NULL && x->right != NULL)
        {

            if (y->left == x)
            {
                temp = x->left;
                y->left = x->left;
                while (temp->right != NULL)
                    temp = temp->right;
                temp->right = x->right;
                x->left = NULL;
                x->right = NULL;
            }
            else
            {
                temp = x->right;
                y->right = x->right;
                while (temp->left != NULL)
                    temp = temp->left;
                temp->left = x->left;
                x->left = NULL;
                x->right = NULL;
            }

            free(x);
            return (p);
        }
        /* this code is for deleting a node with one child*/
        if (x->left == NULL && x->right != NULL)
        {
            if (y->left == x)
                y->left = x->right;
            else
                y->right = x->right;
            x->right = NULL;
            free(x);
            return (p);
        }

        if (x->left != NULL && x->right == NULL)
        {
            if (y->left == x)
                y->left = x->left;
            else
                y->right = x->left;
            x->left = NULL;
            free(x);
            return (p);
        }
        /* this code is for deleting a node with no child*/
        if (x->left == NULL && x->right == NULL)
        {
            if (y->left == x)
                y->left = NULL;
            else
                y->right = NULL;
            free(x);
            return (p);
        }
    }
}

int main()
{
    struct bst *root;
    root = NULL;
    int n, val, num;
    printf("\n enter no. of term:-  ");
    scanf("%d", &n);
    while (n != 0)
    {
        printf("\n enter element:- ");
        scanf("%d", &val);
        root = insert(root, val);
        n--;
    }
    printf("\n display element:-.......");
    inorder(root);
    printf("\n enter element to be deleted:-   ");
    scanf("%d", &num);
    del(root, num);
    printf("\n display element after deleted:-.......");
    inorder(root);
    return 1;
}