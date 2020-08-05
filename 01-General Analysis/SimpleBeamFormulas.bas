Attribute VB_Name = "SimpleBeamFormulas"

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''' BSD 3-Clause License                                                            '''
'''                                                                                 '''
''' Copyright (c) 2020, open-struct-engineer                                        '''
''' All rights reserved.                                                            '''
'''                                                                                 '''
''' Redistribution and use in source and binary forms, with or without              '''
''' modification, are permitted provided that the following conditions are met:     '''
'''                                                                                 '''
''' 1. Redistributions of source code must retain the above copyright notice, this  '''
'''   list of conditions and the following disclaimer.                              '''
'''                                                                                 '''
''' 2. Redistributions in binary form must reproduce the above copyright notice,    '''
'''   this list of conditions and the following disclaimer in the documentation     '''
'''   and/or other materials provided with the distribution.                        '''
'''                                                                                 '''
''' 3. Neither the name of the copyright holder nor the names of its                '''
'''   contributors may be used to endorse or promote products derived from          '''
'''   this software without specific prior written permission.                      '''
'''                                                                                 '''
''' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"     '''
''' AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE       '''
''' IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE  '''
''' DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE    '''
''' FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL      '''
''' DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR      '''
''' SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER      '''
''' CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,   '''
''' OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE   '''
''' OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.            '''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Function point_load(p As Double, a As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''  Function for a point load applied any where on a simply supported beam     '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at Right Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim b As Double     'Distance from load to right support
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim c4 As Double    'Integration constant - C4
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'Right End Fixed End Moment - Clockwise Positive
    

    ''''''''''''''''''''''''''''
    ''  Input Error Handling  ''
    ''''''''''''''''''''''''''''
    If p = 0 Then
        'Capture case where load may be a place holder with no value
        'so return 0 instead of an error
        point_load = 0
    
    Else
    
        If a > L Or a < 0 Then
            point_load = " A , must be between 0 and L"
        
        Else
            ''''''''''''''''''''''''''''
            ''  Common Calculations   ''
            ''''''''''''''''''''''''''''
            
            b = L - a
            
            'Support Reactions
            rl = ((p * b) / L)
            rr = ((p * a) / L)
            
            'Integration Constants
            c4 = ((-1 * rl * a * a * a) / 3) - ((rr * a * a * a) / 3) + ((rr * L * a * a) / 2)
            c2 = (-1 / L) * ((c4) + ((rr * L * L * L) / 3))
            c1 = ((-1 * rr * a * a) / 2) - ((rl * a * a) / 2) + (rr * L * a) + c2
            
        
            '''''''''''''''''''''''''''
            ''  Result Selection     ''
            '''''''''''''''''''''''''''
            
            If result = 0 Then
                
                'Left Support Reaction
                point_load = rl
                
            ElseIf result = 1 Then
                
                'Right Support Reaction
                point_load = rr
            
            ElseIf result = 2 Then
                
                'Shear at x
                If 0 <= x And x <= a And a <> 0 Then
                    vx = rl
                
                ElseIf a < x And x <= L Then
                    vx = -1 * rr
                
                Else
                    vx = 0
                End If
                
                point_load = vx
            
            ElseIf result = 3 Then
            
                'Moment at x
                If 0 <= x And x <= a Then
                    mx = rl * x
                ElseIf a < x And x <= L Then
                    mx = (-1 * rr * x) + (rr * L)
                Else
                    mx = 0
                End If
                
                point_load = mx
                    
            ElseIf result = 4 Then
                
                'Cross Section Rotation/Slope at x
                If 0 <= x And x <= a Then
                    sx = (((rl * x * x) / 2) + c1) / (E * I)
                ElseIf a < x And x <= L Then
                    sx = (((-1 * rr * x * x) / 2) + (rr * L * x) + c2) / (E * I)
                Else
                    sx = 0
                End If
                
                point_load = sx
            
            ElseIf result = 5 Then
                
                'Deflection at x
                If 0 <= x And x <= a Then
                    dx = (((rl * x * x * x) / 6) + (c1 * x)) / (E * I)
                ElseIf a < x And x <= L Then
                    dx = (((-1 * rr * x * x * x) / 6) + ((rr * L * x * x) / 2) + (c2 * x) + c4) / (E * I)
                Else
                    dx = 0
                End If
                
                point_load = dx
            
            ElseIf result = 6 Then
            
                'Fixed End Moment Left
                femL = -1 * (p * a * b * b) / (L * L)
                
                point_load = femL
            
            ElseIf result = 7 Then
                
                'Fixed End Moment Right
                femR = (p * a * a * b) / (L * L)
                point_load = femR
            
            Else
            
                point_load = CVErr(xlErrNA)
                    
            End If
        End If
    End If

End Function

Function point_moment(m As Double, a As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''  Function for a point moment applied any where on a simply supported beam   '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at Right Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim b As Double     'Distance from load to right support
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim c3 As Double    'Integration constant - C3
    Dim c4 As Double    'Integration constant - C4
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'Right End Fixed End Moment - Clockwise Positive
    
    ''''''''''''''''''''''''''''
    ''  Input Error Handling  ''
    ''''''''''''''''''''''''''''
    If m = 0 Then
        'Capture case where load may be a place holder with no value
        'so return 0 instead of an error
        point_moment = 0
    
    Else
    
        If a > L Or a < 0 Then
            point_moment = " A , must be between 0 and L"
        
        Else
            ''''''''''''''''''''''''''''
            ''  Common Calculations   ''
            ''''''''''''''''''''''''''''
            
            b = L - a
            
            'Support Reactions
            rr = m / L
            rl = -1 * rr
            
            'Integration Constants
            c2 = (-1 / L) * ((m * a * a) - (0.5 * m * a * a) + (rl * ((L * L * L) / 6)) + (0.5 * m * L * L))
            c1 = (m * a) + c2
            c3 = 0
            c4 = ((-1 * rl * L * L * L) / 6) - (0.5 * m * L * L) - (c2 * L)
            
        
            '''''''''''''''''''''''''''
            ''  Result Selection     ''
            '''''''''''''''''''''''''''
            
            If result = 0 Then
                
                'Left Support Reaction
                point_moment = rl
                
            ElseIf result = 1 Then
                
                'Right Support Reaction
                point_moment = rr
            
            ElseIf result = 2 Then
                
                'Shear at x
                If 0 <= x And x <= L Then
                    vx = rl
                
                Else
                    vx = 0
                End If
                
                point_moment = vx
            
            ElseIf result = 3 Then
            
                'Moment at x
                If 0 <= x And x <= a Then
                    If x = 0 And a = 0 Then
                        mx = m
                    ElseIf x = L And a = L Then
                        mx = -1 * m
                    Else
                        mx = rl * x
                    End If
                    
                ElseIf a < x And x <= L Then
                    mx = (rl * x) + m
                Else
                    mx = 0
                End If
                
                point_moment = mx
                    
            ElseIf result = 4 Then
                
                'Cross Section Rotation/Slope at x
                If 0 <= x And x <= a Then
                    sx = (((rl * x * x) / 2) + c1) / (E * I)
                ElseIf a < x And x <= L Then
                    sx = ((0.5 * rl * x * x) + (m * x) + c2) / (E * I)
                Else
                    sx = 0
                End If
                
                point_moment = sx
            
            ElseIf result = 5 Then
                
                'Deflection at x
                If 0 <= x And x <= a Then
                    dx = (((1 / 6) * rl * x * x * x) + (c1 * x) + c3) / (E * I)
                ElseIf a < x And x <= L Then
                    dx = (((1 / 6) * rl * x * x * x) + (0.5 * m * x * x) + (c2 * x) + c4) / (E * I)
                Else
                    dx = 0
                End If
                
                point_moment = dx
            
            ElseIf result = 6 Then
            
                'Fixed End Moment Left
                femL = ((-1 * m) / (L * L)) * ((L * L) - (4 * L * a) + (3 * a * a))
                
                point_moment = femL
            
            ElseIf result = 7 Then
                
                'Fixed End Moment Right
                femR = -1 * (m / (L * L)) * ((3 * a * a) - (2 * a * L))
                
                point_moment = femR
            
            Else
            
                point_moment = CVErr(xlErrNA)
                    
            End If
        End If
    End If
    
End Function

Function uniform_load(w As Double, a As Double, b As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Function for a Uniform Distributed Load applied from a to b             '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at Right Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim c As Double     'Width of distributed load
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim c3 As Double    'Integration constant - C3
    Dim c4 As Double    'Integration constant - C4
    Dim c5 As Double    'Integration constant - C5
    Dim c6 As Double    'Integration constant - C6
    Dim c7 As Double    'Integration constant - C7
    Dim c8 As Double    'Integration constant - C8
    Dim c9 As Double    'Integration constant - C9
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'Right End Fixed End Moment - Clockwise Positive
    
    
    ''''''''''''''''''''''''''''
    ''  Input Error Handling  ''
    ''''''''''''''''''''''''''''
    If w = 0 Then
        'Capture case where load may be a place holder with no value
        'so return 0 instead of an error
        uniform_load = 0
    
    Else
    
        If a > L Or b > L Or a = b Or a < 0 Or b < 0 Then
            uniform_load = "A,B cannot be equal and must be between 0 and L"
        ElseIf b < a Then
            uniform_load = "B must be greater than A"
        Else
            ''''''''''''''''''''''''''''
            ''  Common Calculations   ''
            ''''''''''''''''''''''''''''
            
            c = b - a
            
            'Support Reactions
            rl = (w * c) - (((w * c) * (a + (c / 2))) / L)
            rr = (((w * c) * (a + (c / 2))) / L)
            
            'Integration Constants
            c1 = 0
            c2 = ((-1 * w * a * a) / 2)
            c3 = rr * L
            c7 = 0
            c8 = ((-1 * c1 * a * a) / 2) + ((c2 * a * a) / 2) + ((5 * w * a * a * a * a) / 24) + c7
            c9 = ((-1 * rl * b * b * b) / 3) - ((rr * b * b * b) / 3) + ((w * b * b * b * b) / 8) - ((w * a * b * b * b) / 3) - ((c2 * b * b) / 2) + ((c3 * b * b) / 2) + c8
            c6 = ((rr * L * L) / 6) - ((c3 * L) / 2) - (c9 / L)
            c5 = ((-1 * rl * b * b) / 2) + ((w * b * b * b) / 6) - ((w * a * b * b) / 2) - ((rr * b * b) / 2) + (c3 * b) - (c2 * b) + c6
            c4 = ((w * a * a * a) / 3) + (c2 * a) + c5 - (c1 * a)
            
            '''''''''''''''''''''''''''
            ''  Result Selection     ''
            '''''''''''''''''''''''''''
            
            If result = 0 Then
                
                'Left Support Reaction
                uniform_load = rl
                
            ElseIf result = 1 Then
                
                'Right Support Reaction
                uniform_load = rr
            
            ElseIf result = 2 Then
                
                'Shear at x
                If 0 <= x And x <= a Then
                    vx = rl
                
                ElseIf a < x And x <= b Then
                    vx = rl - (w * (x - a))
                
                ElseIf b < x And x <= L Then
                    vx = -1 * rr
                
                Else
                    vx = 0
                    
                End If
                
                uniform_load = vx
            
            ElseIf result = 3 Then
            
                'Moment at x
                If 0 <= x And x <= a Then
                    mx = (rl * x) + c1
                
                ElseIf a < x And x <= b Then
                    mx = (rl * x) - ((w * x * x) / 2) + (w * a * x) + c2
                
                ElseIf b < x And x <= L Then
                    mx = (-1 * rr * x) + c3
                
                Else
                    mx = 0
                    
                End If
                
                uniform_load = mx
                    
            ElseIf result = 4 Then
                
                'Cross Section Rotation/Slope at x
                If 0 <= x And x <= a Then
                    sx = (((rl * x * x) / 2) + (c1 * x) + c4) / (E * I)
                
                ElseIf a < x And x <= b Then
                    sx = (((rl * x * x) / 2) - ((w * x * x * x) / 6) + ((w * a * x * x) / 2) + (c2 * x) + c5) / (E * I)
                
                ElseIf b < x And x <= L Then
                    sx = (((-1 * rr * x * x) / 2) + (c3 * x) + c6) / (E * I)
                
                Else
                    sx = 0
                    
                End If
                
                uniform_load = sx
            
            ElseIf result = 5 Then
                
                'Deflection at x
                 If 0 <= x And x <= a Then
                    dx = (((rl * x * x * x) / 6) + ((c1 * x * x) / 2) + (c4 * x) + c7) / (E * I)
                
                ElseIf a < x And x <= b Then
                    dx = (((rl * x * x * x) / 6) - ((w * x * x * x * x) / 24) + ((w * a * x * x * x) / 6) + ((c2 * x * x) / 2) + (c5 * x) + c8) / (E * I)
                
                ElseIf b < x And x <= L Then
                    dx = (((-1 * rr * x * x * x) / 6) + ((c3 * x * x) / 2) + (c6 * x) + c9) / (E * I)
                
                Else
                    dx = 0
                    
                End If
                
                uniform_load = dx
            
            ElseIf result = 6 Then
            
                'Fixed End Moment Left
                femL = ((rr * L * L * 0.5) - (c3 * L) - c6 - (2 * c4)) / (-0.5 * L)
                
                uniform_load = femL
            
            ElseIf result = 7 Then
                
                'Fixed End Moment Right
                femR = ((-1 * c4) + ((((rr * L * L * 0.5) - (c3 * L) - c6 - (2 * c4)) / (-0.5 * L)) * (L / 3))) * (6 / L)
                
                uniform_load = femR
            
            Else
            
                uniform_load = CVErr(xlErrNA)
                    
            End If
        End If
    End If
    
End Function

Function variable_load(w1 As Double, w2 As Double, a As Double, b As Double, L As Double, E As Double, I As Double, x As Double, result As Integer) As Variant

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Function for a Variable Distributed Load applied from a to b            '''
    '''                                                                             '''
    '''     Important Note:                                                         '''
    '''     All inputs must have consistent units                                   '''
    '''                                                                             '''
    '''     Result key:                                                             '''
    '''     0 = Left Reaction                                                       '''
    '''     1 = Right Reaction                                                      '''
    '''     2 = Shear at x                                                          '''
    '''     3 = Moment at x                                                         '''
    '''     4 = Cross Section Rotation/Slope at x                                   '''
    '''     5 = Deflection at x                                                     '''
    '''     6 = Fixed End Moment at Left Support (clockwise positive)               '''
    '''     7 = Fixed End Moment at Right Support (clockwise positive)              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''                                                                             '''
    '''     Sign Convention:                                                        '''
    '''     Loads applied in the (-)y direction are positive                        '''
    '''     Clockwise moments are positive                                          '''
    '''                                                                             '''
    '''     Reactions in the (+)y direction are positive                            '''
    '''                                                                             '''
    '''     Internal:                                                               '''
    '''     Shear is positive in the (+)y direction                                 '''
    '''     Moment is positive clockwise                                            '''
    '''     Cross Section Rotation/Slope is positive counter-clockwise              '''
    '''     Upward deflection is in the (+)y direction                              '''
    '''                                                                             '''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    ''''''''''''''''''''''''''''''
    ''  Variable Definitions    ''
    ''''''''''''''''''''''''''''''
    
    Dim c As Double     'Width of distributed load
    Dim s As Double     'Slope of load
    Dim xbar As Double  'Centroid of load
    Dim w As Double     'Area of Load or Equivalent Point Load
    Dim rl As Double    'Left Reaction for a simple span beam
    Dim rr As Double    'Right Reaction foa smiple span beam
    Dim c1 As Double    'Integration constant - C1
    Dim c2 As Double    'Integration constant - C2
    Dim c3 As Double    'Integration constant - C3
    Dim c4 As Double    'Integration constant - C4
    Dim c5 As Double    'Integration constant - C5
    Dim c6 As Double    'Integration constant - C6
    Dim c7 As Double    'Integration constant - C7
    Dim c8 As Double    'Integration constant - C8
    Dim c9 As Double    'Integration constant - C9
    Dim vx As Double    'Shear at x
    Dim mx As Double    'Moment at x
    Dim sx As Double    'Cross section Rotation/Slope at x
    Dim dx As Double    'Deflection at x
    Dim femL As Double  'Left End Fixed End Moment - Clockwise Positive
    Dim femR As Double  'Right End Fixed End Moment - Clockwise Positive
    
    
    ''''''''''''''''''''''''''''
    ''  Input Error Handling  ''
    ''''''''''''''''''''''''''''
    If w1 = 0 And w2 = 0 Then
        'Capture case where load may be a place holder with no value
        'so return 0 instead of an error
        variable_load = 0
    
    Else
    
        If a > L Or b > L Or a = b Or a < 0 Or b < 0 Then
            variable_load = "A,B cannot be equal and must be between 0 and L"
        ElseIf b < a Then
            variable_load = "B must be greater than A"
        ElseIf Sgn(w1) <> Sgn(w2) And w1 <> 0 And w2 <> 0 Then
            variable_load = "W1,W2 must have the same sign"
        Else
            ''''''''''''''''''''''''''''
            ''  Common Calculations   ''
            ''''''''''''''''''''''''''''
            
            c = b - a
            s = (w2 - w1) / c
            xbar = (c * ((2 * w2) + w1)) / (3 * (w2 + w1))
            w = c * ((w1 + w2) / 2)
            
            'Support Reactions
            rr = (w * (a + xbar)) / L
            rl = w - rr
            
            'Integration Constants
            c1 = 0
            c2 = c1 + ((a * a * a * s) / 6) + ((a * a * (w1 - (s * a))) / 2) + ((((s * a) - (2 * w1)) * a * a) / 2)
            c3 = rr * L
            c7 = 0
            c8 = ((-1 * c1 * a * a) / 2) - ((a * a * a * a * a * s) / 30) - ((a * a * a * a * (w1 - (s * a))) / 8) - ((((s * a) - (2 * w1)) * a * a * a * a) / 6) + ((c2 * a * a) / 2) + c7
            c9 = ((-1 * rl * b * b * b) / 3) + ((b * b * b * b * b * s) / 30) + ((b * b * b * b * (w1 - (s * a))) / 8) + ((((s * a) - (2 * w1)) * a * b * b * b) / 6) - ((c2 * b * b) / 2) + c8 - ((rr * b * b * b) / 3) + ((c3 * b * b) / 2)
            c6 = (((rr * L * L * L) / 6) - ((c3 * L * L) / 2) - c9) / L
            c5 = ((-1 * rr * b * b) / 2) + (c3 * b) + c6 - ((rl * b * b) / 2) + ((b * b * b * b * s) / 24) + ((b * b * b * (w1 - (s * a))) / 6) + ((((s * a) - (2 * w1)) * a * b * b) / 4) - (c2 * b)
            c4 = ((-1 * a * a * a * a * s) / 24) - ((a * a * a * (w1 - (s * a))) / 6) - ((((s * a) - (2 * w1)) * a * a * a) / 4) + (c2 * a) + c5 - (c1 * a)
            
            '''''''''''''''''''''''''''
            ''  Result Selection     ''
            '''''''''''''''''''''''''''
            
            If result = 0 Then
                
                'Left Support Reaction
                variable_load = rl
                
            ElseIf result = 1 Then
                
                'Right Support Reaction
                variable_load = rr
            
            ElseIf result = 2 Then
                
                'Shear at x
                If 0 <= x And x <= a Then
                    vx = rl
                
                ElseIf a < x And x <= b Then
                    vx = rl - ((x * x * s) / 2) - (x * (w1 - (s * a))) - ((((s * a) - (2 * w1)) * a) / 2)
                
                ElseIf b < x And x <= L Then
                    vx = -1 * rr
                
                Else
                    vx = 0
                    
                End If
                
                variable_load = vx
            
            ElseIf result = 3 Then
            
                'Moment at x
                If 0 <= x And x <= a Then
                    mx = (rl * x) + c1
                
                ElseIf a < x And x <= b Then
                    mx = (rl * x) - ((x * x * x * s) / 6) - ((x * x * (w1 - (s * a))) / 2) - ((((s * a) - (2 * w1)) * a * x) / 2) + c2
                
                ElseIf b < x And x <= L Then
                    mx = (-1 * rr * x) + c3
                
                Else
                    mx = 0
                    
                End If
                
                variable_load = mx
                    
            ElseIf result = 4 Then
                
                'Cross Section Rotation/Slope at x
                If 0 <= x And x <= a Then
                    sx = (((rl * x * x) / 2) + (c1 * x) + c4) / (E * I)
                
                ElseIf a < x And x <= b Then
                    sx = (((rl * x * x) / 2) - ((x * x * x * x * s) / 24) - ((x * x * x * (w1 - (s * a))) / 6) - ((((s * a) - (2 * w1)) * a * x * x) / 4) + (c2 * x) + c5) / (E * I)
                
                ElseIf b < x And x <= L Then
                    sx = (((-1 * rr * x * x) / 2) + (c3 * x) + c6) / (E * I)
                
                Else
                    sx = 0
                    
                End If
                
                variable_load = sx
            
            ElseIf result = 5 Then
                
                'Deflection at x
                 If 0 <= x And x <= a Then
                    dx = (((rl * x * x * x) / 6) + ((c1 * x * x) / 2) + (c4 * x) + c7) / (E * I)
                
                ElseIf a < x And x <= b Then
                    dx = (((rl * x * x * x) / 6) - ((x * x * x * x * x * s) / 120) - ((x * x * x * x * (w1 - (s * a))) / 24) - ((((s * a) - (2 * w1)) * a * x * x * x) / 12) + ((c2 * x * x) / 2) + (c5 * x) + c8) / (E * I)
                
                ElseIf b < x And x <= L Then
                    dx = (((-1 * rr * x * x * x) / 6) + ((c3 * x * x) / 2) + (c6 * x) + c9) / (E * I)
                
                Else
                    dx = 0
                    
                End If
                
                variable_load = dx
            
            ElseIf result = 6 Then
            
                'Fixed End Moment Left
                femL = ((rr * L * L * 0.5) - (c3 * L) - c6 - (2 * c4)) / (-0.5 * L)
                
                variable_load = femL
            
            ElseIf result = 7 Then
                
                'Fixed End Moment Right
                femR = ((-1 * c4) + ((((rr * L * L * 0.5) - (c3 * L) - c6 - (2 * c4)) / (-0.5 * L)) * (L / 3))) * (6 / L)
                
                variable_load = femR
            
            Else
            
                variable_load = CVErr(xlErrNA)
                    
            End If
        End If
    End If
End Function
