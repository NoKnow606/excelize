// Copyright 2016 - 2025 The excelize Authors. All rights reserved. Use of
// this source code is governed by a BSD-style license that can be found in
// the LICENSE file.

package excelize

import (
	"container/list"
	"sync"
)

// lruCache implements a thread-safe LRU (Least Recently Used) cache
// with a maximum size limit. When the cache is full, the least recently
// used item is evicted to make room for new items.
type lruCache struct {
	mu       sync.RWMutex
	capacity int
	cache    map[string]*list.Element
	lruList  *list.List
}

// lruEntry represents a key-value pair in the LRU cache
type lruEntry struct {
	key   string
	value interface{}
}

// newLRUCache creates a new LRU cache with the specified capacity
func newLRUCache(capacity int) *lruCache {
	return &lruCache{
		capacity: capacity,
		cache:    make(map[string]*list.Element),
		lruList:  list.New(),
	}
}

// Load retrieves a value from the cache. Returns (value, true) if found,
// (nil, false) if not found. Moves the accessed item to the front (most recent).
func (c *lruCache) Load(key string) (interface{}, bool) {
	c.mu.Lock()
	defer c.mu.Unlock()

	if elem, ok := c.cache[key]; ok {
		// Move to front (most recently used)
		c.lruList.MoveToFront(elem)
		return elem.Value.(*lruEntry).value, true
	}
	return nil, false
}

// Store adds or updates a value in the cache. If the cache is at capacity,
// the least recently used item is evicted. Returns true if an item was evicted.
func (c *lruCache) Store(key string, value interface{}) bool {
	c.mu.Lock()
	defer c.mu.Unlock()

	// If key exists, update and move to front
	if elem, ok := c.cache[key]; ok {
		c.lruList.MoveToFront(elem)
		elem.Value.(*lruEntry).value = value
		return false
	}

	// Check capacity and evict if necessary
	evicted := false
	if c.lruList.Len() >= c.capacity {
		// Remove least recently used (back of list)
		oldest := c.lruList.Back()
		if oldest != nil {
			c.lruList.Remove(oldest)
			oldKey := oldest.Value.(*lruEntry).key
			delete(c.cache, oldKey)
			evicted = true
		}
	}

	// Add new entry to front
	entry := &lruEntry{key: key, value: value}
	elem := c.lruList.PushFront(entry)
	c.cache[key] = elem

	return evicted
}

// Clear removes all items from the cache
func (c *lruCache) Clear() {
	c.mu.Lock()
	defer c.mu.Unlock()

	c.cache = make(map[string]*list.Element)
	c.lruList = list.New()
}

// Len returns the current number of items in the cache
func (c *lruCache) Len() int {
	c.mu.RLock()
	defer c.mu.RUnlock()

	return c.lruList.Len()
}

// Range calls f for each key-value pair in the cache.
// If f returns false, iteration stops.
func (c *lruCache) Range(f func(key string, value interface{}) bool) {
	c.mu.RLock()
	defer c.mu.RUnlock()

	for elem := c.lruList.Front(); elem != nil; elem = elem.Next() {
		entry := elem.Value.(*lruEntry)
		if !f(entry.key, entry.value) {
			break
		}
	}
}

// Delete removes a key from the cache. Returns true if the key was present.
func (c *lruCache) Delete(key string) bool {
	c.mu.Lock()
	defer c.mu.Unlock()

	if elem, ok := c.cache[key]; ok {
		c.lruList.Remove(elem)
		delete(c.cache, key)
		return true
	}
	return false
}
