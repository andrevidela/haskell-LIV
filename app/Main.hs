{-# LANGUAGE NoImplicitPrelude #-}
{-# LANGUAGE OverloadedStrings #-}
module Main where

import Lib
import Codec.Xlsx
import Control.Lens
import qualified Data.ByteString.Lazy as L
import Data.Maybe
import Protolude

filename :: Text
filename = "LIV-222741+1_2.xlsx"

newtype M_Volt = M_Volt Double deriving (Eq, Show)
newtype Volt = Volt Double deriving (Eq, Show)
newtype M_Apcm2 = M_Apcm2 Double deriving (Eq, Show)
newtype Apcm2 = Apcm2 Double deriving (Eq, Show)
newtype Ampere = Ampere Double deriving (Eq, Show)
newtype M_Ampere = M_Ampere Double deriving (Eq, Show)
newtype Wpm2 = Wpm2 Double deriving (Eq, Show)
newtype OhmCm2 = OhmCm2 Double deriving (Eq, Show)

fromMiliAmp :: M_Ampere -> Ampere
fromMiliAmp (M_Ampere value) = Ampere $ value / 1000

fromMiliVolt :: M_Volt -> Volt
fromMiliVolt (M_Volt value) = Volt $ value / 1000

fromMiliAmpPerCm2 :: M_Apcm2 -> Apcm2
fromMiliAmpPerCm2 (M_Apcm2 v) = Apcm2 $ v / 1000

data SheetHeader = SheetHeader
    { name :: Text
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
    } deriving (Eq, Show)

-- SI Projections for SheetHeader
vocSI ::SheetHeader -> Volt
vocSI = fromMiliVolt . voc

jscSI :: SheetHeader -> Apcm2
jscSI = fromMiliAmpPerCm2 . jsc

vmppSI :: SheetHeader -> Volt
vmppSI = fromMiliVolt . vmpp

jmppSI :: SheetHeader -> Apcm2
jmppSI = fromMiliAmpPerCm2 . jmpp

iscSI :: SheetHeader -> Ampere
iscSI = fromMiliAmp . isc

getFile :: IO Xlsx
getFile = do bs <- L.readFile $ toS filename
             return $ toXlsx bs

tryCellDouble :: CellValue -> Maybe Double
tryCellDouble (CellDouble double) = Just double
tryCellDouble _                   = Nothing

tryCellText :: CellValue -> Maybe Text
tryCellText (CellText text) = Just text
tryCellText _               = Nothing

tryConvertCell :: (Double -> a) -> CellValue -> Maybe a
tryConvertCell constructor cell = do d <- tryCellDouble cell
                                     pure $ constructor d

getHeaderInfo :: Worksheet -> Maybe SheetHeader
getHeaderInfo sheet = do name <- getCellNr 1 >>= tryCellText
                         eff  <- getCellNr 2 >>= tryCellDouble
                         voc  <- getCellNr 3 >>= tryConvertCell M_Volt
                         jsc  <- getCellNr 4 >>= tryConvertCell M_Apcm2
                         ff   <- getCellNr 5 >>= tryCellDouble
                         vmpp <- getCellNr 6 >>= tryConvertCell M_Volt
                         jmpp <- getCellNr 7 >>= tryConvertCell M_Apcm2
                         pmmp <- getCellNr 8 >>= tryConvertCell Wpm2
                         rsc  <- getCellNr 9 >>= tryConvertCell OhmCm2
                         roc  <- getCellNr 10 >>= tryConvertCell OhmCm2
                         isc  <- getCellNr 11 >>= tryConvertCell M_Ampere
                         pure SheetHeader
                             { name = name
                             , eff = eff
                             , voc = voc
                             , jsc = jsc
                             , ff = ff
                             , vmpp = vmpp
                             , jmpp = jmpp
                             , pmmp = pmmp
                             , rsc = rsc
                             , roc = roc
                             , isc = isc
                             }
    where
        getCellNr :: Int -> Maybe CellValue
        getCellNr col = sheet ^? ixCell (35, col) . cellValue . _Just

getHeaderInfo' :: Worksheet -> [Maybe CellValue]
getHeaderInfo' sheet = map getCellNr [1,2,3,4,5,6,7,8,9,10,11]
  where
      getCellNr :: Int -> Maybe CellValue
      getCellNr col = sheet ^? ixCell (34,col) . cellValue . _Just

getHeaders :: Xlsx -> [SheetHeader]
getHeaders doc = mapMaybe (getHeaderInfo . snd) (_xlSheets doc)

main :: IO ()
main = do file <- getFile
          let headers = getHeaders file
          mapM_ print headers
          
