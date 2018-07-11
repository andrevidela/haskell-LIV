module Main where

import Lib
import Codec.Xlsx
import Control.Lens
import qualified Data.ByteString.Lazy as L

filename :: String
filename = "LIV-222741+1.xls"

newtype M_Volt = M_Volt Double
newtype Volt = Volt Double
newtype M_Apcm2 = M_Apcm2 Double
newtype Ampere = Ampere Double
newtype M_Ampere = M_Ampere Double
newtype Wpm2 = Wpm2 Double
newtype OhmCm2 = OhmCm2 Double

fromMiliAmp :: M_Ampere -> Ampere
fromMiliAmp (M_Ampere value) = Ampere $ value / 1000

fromMiliVolt :: M_Volt -> Volt
fromMiliVolt (M_Volt value) = Volt $ value / 1000

data SheetHeader = SheetHeader
    { name :: String
    , eff :: Double
    , voc :: M_Volt
    , jsc :: M_Apcm2
    , ff :: Double
    , vmpp :: M_Volt
    , jmpp :: M_Apcm2
    , pmmp :: Wpm2
    , rsc :: OhmCm2 
    , roc :: OhmCm2
    , isc :: M_Ampere
    }

getFile :: IO Xlsx
getFile = do bs <- L.readFile filename
             return $ toXlsx bs

main :: IO ()
main = do file <- getFile
          return ()
          
