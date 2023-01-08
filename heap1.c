// Heap
// Binary Heap can be visualized array as a complete binary tree
// Arr[0] element will be treated as root
// length(A) – size of array
// heapSize(A) – size of heap
// Generally used when we are dealing with minimum and maximum elements
// For ith node
// (i-1)/2	Parent
// (2*i)+1	Left child
// (2*i)+2	Right Child
// ----------------------------------------------------------------Advantages
// Can be of 2 types: min heap and max heap
// Min heap keeps smallest and element and top and max keeps the largest
// O(1) for dealing with min or max elements
// ----------------------------------------------------------------Disadvantages
// Random access not possible
// Only min or max element is available for accessibility
// Applications

// Suitable for applications dealing with priority
// Scheduling algorithm
// caching

#include <stdio.h>
int size = 0;
void swap(int *p, int *q)
{
    int temp = *q;
    *q = *p;
    *p = temp;
}
void heapify(int array[], int size, int i)
{
    if (size == 1)
    {
        printf("Single element in the heap");
    }
    else
    {
        int largest = i;
        int l = 2 * i + 1;
        int r = 2 * i + 2;
        if (l < size && array[l] > array[largest])
            largest = l;
        if (r < size && array[r] > array[largest])
            largest = r;
        if (largest != i)
        {
            swap(&array[i], &array[largest]);
            heapify(array, size, largest);
        }
    }
}
void insert(int array[], int newNumber)
{
    if (size == 0)
    {
        array[0] = newNumber;
        size += 1;
    }
    else
    {
        array[size] = newNumber;
        size += 1;
        for (int i = size / 2 - 1; i >= 0; i--)
        {
            heapify(array, size, i);
        }
    }
}
void deleteRoot(int array[], int num)
{
    int i;
    for (i = 0; i < size; i++)
    {
        if (num == array[i])
            break;
    }

    swap(&array[i], &array[size - 1]);
    size -= 1;
    for (int i = size / 2 - 1; i >= 0; i--)
    {
        heapify(array, size, i);
    }
}
void printArray(int array[], int size)
{
    for (int i = 0; i < size; ++i)
        printf("%d ", array[i]);
    printf("\n");
}
int main()
{
    int array[10];

    insert(array, 4);
    insert(array, 2);
    insert(array, 8);
    insert(array, 1);
    insert(array, 3);
    insert(array, 6);
    printf("Max Heap array: ");
    printArray(array, size);

    deleteRoot(array, 4);

    printf("After deleting an element: ");

    printArray(array, size);
}