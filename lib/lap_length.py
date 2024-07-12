def lap_legth(dia,p1,n1):
    #Cd = 
    n1 = 0.7
    n2 = (1 if dia <= 32 else (132-dia)/100)
    fctk = 2.2
    fctd = fctk/1.5
    fdb = 2.25*n1*n2*fctd
    sigma_sd = 500/1.15
    fyd = 500   
    alpha1 = 1
    alpha2 = 1#-0.15*(Cd-dia)/dia
    alpha3 = 1#*3.41*(dia/2)**2*(sigma_sd/fyd)
    alpha5 = 1 
    alpha6 = (((p1/25)**0.5) if ((p1/25)**0.5)>1 and ((p1/25)**0.5)<1.5 else (1 if ((p1/25)**0.5)<1 else 1.5)) 
    lb_rqd = (dia/4)*(sigma_sd/fdb)

    l0 = alpha1 * alpha2 * alpha3 * alpha5 * alpha6 * lb_rqd
    l0min = max(0.3*alpha6*lb_rqd, 15*dia, 200)

    print("n1 = ", n1)
    print("n2 = ", n2)
    print("fctd = ", fctd)
    print("fbd = ", fdb)
    print("sigma_sd = ", sigma_sd)
    print("fyd = ", fyd)
    print("alpha1 = ", alpha1)
    print("alpha2 = ", alpha2)
    print("alpha3 = ", alpha3)
    print("alpha5 = ", alpha5)
    print("alpha6 = ", alpha6)
    print("lb_rqd = ", lb_rqd)
    print("l0 = ", l0)
    print("l0min = ", l0min)
    print("#"*50)    
    print(l0/dia)
    if l0 < l0min:
        l0 = l0min  
        print("Lap length is less than minimum lap length, hence lap length is taken as minimum lap length")
    return l0

diaList = [ 16, 20, 25, 32, ]

for dia in diaList:

    print(lap_legth(dia, 30))
    print(dia)
    print("#"*50)
