'use client'
import { useRef, useEffect } from 'react'
import { Provider } from 'react-redux'
import { makeStore } from '../lib/store'
import { useRouter } from 'next/navigation'

export default function StoreProvider({ children }) {
  const storeRef = useRef(undefined)
  if (!storeRef.current) {
    // Create the store instance the first time this renders
    storeRef.current = makeStore()
  }

  const router = useRouter()
  useEffect(() => {
    if (typeof window !== 'undefined') {
      window.nextRouter = router
    }
  }, [router])

  return <Provider store={storeRef.current}>{children}</Provider>
}